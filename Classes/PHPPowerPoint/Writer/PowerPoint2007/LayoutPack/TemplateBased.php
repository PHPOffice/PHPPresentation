<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Writer_PowerPoint2007_LayoutPack_TemplateBased
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_LayoutPack_TemplateBased extends PHPPowerPoint_Writer_PowerPoint2007_LayoutPack
{
	/**
	 * PHPPowerPoint_Writer_PowerPoint2007_LayoutPack_TemplateBased
	 */
	public function __construct($fileName = '') {
		// Check if file exists
		if (!file_exists($fileName)) {
			throw new Exception("Could not open " . $fileName . " for reading! File does not exist.");
		}

		// Master slide relations
		$this->_masterSlideRelations = array();

		// Theme relations
		$this->_themeRelations = array();

		// Layout relations
		$this->_layoutRelations = array();

		// Open package
		$package = new ZipArchive;
		$package->open($fileName);

	    // Read relations and search for officeDocument
	    $layoutId = -1;
        $relations = simplexml_load_string($package->getFromName("_rels/.rels"));
        foreach ($relations->Relationship as $rel) {
            if ($rel["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument") {
                // Found office document! Search for master slide...
                $presentationRelations = simplexml_load_string($package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/_rels/" . basename($rel["Target"]) . ".rels")) );
                foreach ($presentationRelations->Relationship as $presRel) {
                    if ($presRel["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster") {
                        // Found slide master!
                        $slideMasterId = str_replace('slideMaster', '', basename($presRel["Target"], '.xml'));
                        $this->_masterSlides[] = array('masterid' => $slideMasterId, 'body' => $package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/" . basename($presRel["Target"])) ));
                        
                        // Search for theme & slide layouts
                        $masterRelations = simplexml_load_string($package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/_rels/" . basename($presRel["Target"]) . ".rels")) );
                        foreach ($masterRelations->Relationship as $masterRel) {
                            if ($masterRel["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                                // Found theme!
                                $themeId = str_replace('theme', '', basename($masterRel["Target"], '.xml'));
                                $this->_themes[$themeId - 1] = array('masterid' => $slideMasterId, 'body' => $package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/" . dirname($masterRel["Target"]) . "/" . basename($masterRel["Target"])) ));

                                // Search for theme relations
                                $themeRelations = @simplexml_load_string($package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/" . dirname($masterRel["Target"]) .  "/_rels/" . basename($masterRel["Target"]) . ".rels")) );
                                if ($themeRelations && $themeRelations->Relationship) {
                                    foreach ($themeRelations->Relationship as $themeRel) {
                                    	if ($themeRel["Type"] != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
                                    		&& $themeRel["Type"] != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
                                    		&& $themeRel["Type"] != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                                        	// Theme relation
                                        	$this->_themeRelations[] = array(
                                        		'masterid'     => $slideMasterId, 
                                        		'id'           => $themeRel["Id"],
                                        		'type'         => $themeRel["Type"],
                                        		'contentType'  => '',
                                        		'target'       => $themeRel["Target"],
                                        		'contents'     => $package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/" . dirname($masterRel["Target"]) . "/" . dirname($themeRel["Target"]) . "/" . basename($themeRel["Target"])) )
                                        	);
                                    	}
                                    }
                                }
                            } else if ($masterRel["Type"] == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout") {
                            	// Found slide layout!
                                $layoutId = str_replace('slideLayout', '', basename($masterRel["Target"], '.xml'));
                                $layout = array(
											'masterid'  => $slideMasterId,
                                			'name' 		=> '-unknown-',
                                			'body' 		=> $package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/" . dirname($masterRel["Target"]) . "/" . basename($masterRel["Target"])) )
                                );
                                $layoutXml = null;
                                if(utf8_encode(utf8_decode($layout['body'])) == $layout['body'])
                                	$layoutXml = simplexml_load_string($layout['body']);
                                else
                                	$layoutXml = simplexml_load_string(utf8_encode($layout['body']));
            					$layoutXml->registerXPathNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            					$slide = $layoutXml->xpath('/p:sldLayout/p:cSld');
            					$layout['name'] = (string)$slide[0]['name'];
                                $this->_layouts[$layoutId] = $layout;

                                // Search for slide layout relations
                                $layoutRelations = @simplexml_load_string($package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/" . dirname($masterRel["Target"]) .  "/_rels/" . basename($masterRel["Target"]) . ".rels")) );
                                if ($layoutRelations && $layoutRelations->Relationship) {
                                    foreach ($layoutRelations->Relationship as $layoutRel) {
                                        if ($layoutRel["Type"] != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
                                        		&& $layoutRel["Type"] != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
                                        		&& $layoutRel["Type"] != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                                        	// Layout relation
                                        	$this->_layoutRelations[] = array(
                                        		'layoutId'     => $layoutId,
                                        		'id'           => $layoutRel["Id"],
                                        		'type'         => $layoutRel["Type"],
                                        		'contentType'  => '',
                                        		'target'       => $layoutRel["Target"],
                                        		'contents'     => $package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/" . dirname($masterRel["Target"]) . "/" . dirname($layoutRel["Target"]) . "/" . basename($layoutRel["Target"])) )
                                        	);
                                        }
                                    }
                                }
                            } else {
                            	// Master slide relation
                            	$this->_masterSlideRelations[] = array(
									'masterid'     => $slideMasterId,
                            		'id'           => $masterRel["Id"],
                            		'type'         => $masterRel["Type"],
                            		'contentType'  => '',
                            		'target'       => $masterRel["Target"],
                            		'contents'     => $package->getFromName( $this->absoluteZipPath(dirname($rel["Target"]) . "/" . dirname($presRel["Target"]) . "/" . dirname($masterRel["Target"]) . "/" . basename($masterRel["Target"])) )
                            	);
                            }
                        }
                    }
                }

                break;
            }
        }

        // Sort master slides
        usort($this->_masterSlides, array("PHPPowerPoint_Writer_PowerPoint2007_LayoutPack_TemplateBased", "cmp_master"));
        
		// Close package
		$package->close();
	}
	
	/**
	 * Compare master slides
	 * 
	 * @param array $a
	 * @param array $b
	 */
	public static function cmp_master($a, $b) {
		if ($a['masterid'] == $b['masterid']) {
	        return 0;
	    }
	    return ($a['masterid'] < $b['masterid']) ? -1 : 1;	
	}

    /**
     * Determine absolute zip path
     *
     * @param string $path
     * @return string
     */
    protected function absoluteZipPath($path) {
        $path = str_replace(array('/', '\\'), DIRECTORY_SEPARATOR, $path);
        $parts = array_filter(explode(DIRECTORY_SEPARATOR, $path), 'strlen');
        $absolutes = array();
        foreach ($parts as $part) {
            if ('.' == $part) continue;
            if ('..' == $part) {
                array_pop($absolutes);
            } else {
                $absolutes[] = $part;
            }
        }
        return implode('/', $absolutes);
    }
}
