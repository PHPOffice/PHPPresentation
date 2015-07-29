<?php
/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Reader;

use ZipArchive;
use PhpOffice\Common\XMLReader;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;

/**
 * Serialized format reader
 */
class PowerPoint2007 implements ReaderInterface
{
    /**
     * Output Object
     * @var PhpPresentation
     */
    protected $oPhpPresentation;
    /**
     * Output Object
     * @var \ZipArchive
     */
    protected $oZip;
    /**
     * @var string[]
     */
    protected $arrayRels = array();

    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     *
     * @param  string $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function canRead($pFilename)
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     *
     * @param  string    $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function fileSupportsUnserializePhpPresentation($pFilename = '')
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }
        
        $oZip = new ZipArchive();
        // Is it a zip ?
        if ($oZip->open($pFilename) === true) {
            // Is it an OpenXML Document ?
            // Is it a Presentation ?
            if (is_array($oZip->statName('[Content_Types].xml')) && is_array($oZip->statName('ppt/presentation.xml'))) {
                return true;
            }
        }
        return false;
    }

    /**
     * Loads PhpPresentation Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     * @throws \Exception
     */
    public function load($pFilename)
    {
        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPresentation\Reader\PowerPoint2007: " . $pFilename . ".");
        }

        return $this->loadFile($pFilename);
    }

    /**
     * Load PhpPresentation Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     */
    protected function loadFile($pFilename)
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();
        
        $this->oZip = new ZipArchive();
        $this->oZip->open($pFilename);
        $docPropsCore = $this->oZip->getFromName('docProps/core.xml');
        if ($docPropsCore !== false) {
            $this->loadDocumentProperties($docPropsCore);
        }
        
        $pptPresentation = $this->oZip->getFromName('ppt/presentation.xml');
        if ($pptPresentation !== false) {
            $this->loadSlides($pptPresentation);
        }

        return $this->oPhpPresentation;
    }
    
    /**
     * Read Document Properties
     * @param string $sPart
     */
    protected function loadDocumentProperties($sPart)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            $arrayProperties = array(
                    '/cp:coreProperties/dc:creator' => 'setCreator',
                    '/cp:coreProperties/cp:lastModifiedBy' => 'setLastModifiedBy',
                    '/cp:coreProperties/dc:title' => 'setTitle',
                    '/cp:coreProperties/dc:description' => 'setDescription',
                    '/cp:coreProperties/dc:subject' => 'setSubject',
                    '/cp:coreProperties/cp:keywords' => 'setKeywords',
                    '/cp:coreProperties/cp:category' => 'setCategory',
                    '/cp:coreProperties/dcterms:created' => 'setCreated',
                    '/cp:coreProperties/dcterms:modified' => 'setModified',
            );
            $oProperties = $this->oPhpPresentation->getProperties();
            foreach ($arrayProperties as $path => $property) {
                if (is_object($oElement = $xmlReader->getElement($path))) {
                    if ($oElement->hasAttribute('xsi:type') && $oElement->getAttribute('xsi:type') == 'dcterms:W3CDTF') {
                        $oDateTime = new \DateTime();
                        $oDateTime->createFromFormat(\DateTime::W3C, $oElement->nodeValue);
                        $oProperties->{$property}($oDateTime->getTimestamp());
                    } else {
                        $oProperties->{$property}($oElement->nodeValue);
                    }
                }
            }
        }
    }
    
    /**
     * Extract all slides
     */
    protected function loadSlides($sPart)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            $fileRels = 'ppt/_rels/presentation.xml.rels';
            $this->loadRels($fileRels);
            foreach ($xmlReader->getElements('/p:presentation/p:sldIdLst/p:sldId') as $oElement) {
                $rId = $oElement->getAttribute('r:id');
                $pathSlide = isset($this->arrayRels[$fileRels][$rId]) ? $this->arrayRels[$fileRels][$rId] : '';
                if (!empty($pathSlide)) {
                    $pptSlide = $this->oZip->getFromName('ppt/'.$pathSlide);
                    if ($pptSlide !== false) {
                        $this->loadRels('ppt/slides/_rels/'.basename($pathSlide).'.rels');
                        $this->loadSlide($pptSlide, basename($pathSlide));
                    }
                }
            }
        }
    }

    /**
     * Extract data from slide
     * @param string $sPart
     * @param string $baseFile
     */
    protected function loadSlide($sPart, $baseFile)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            // Core
            $this->oPhpPresentation->createSlide();
            $this->oPhpPresentation->setActiveSlideIndex($this->oPhpPresentation->getSlideCount() - 1);
            foreach ($xmlReader->getElements('/p:sld/p:cSld/p:spTree/*') as $oNode) {
                switch ($oNode->tagName) {
                    case 'p:pic':
                        $this->loadShapeDrawing($xmlReader, $oNode, $baseFile);
                        break;
                    case 'p:sp':
                        $this->loadShapeRichText($xmlReader, $oNode, $baseFile);
                        break;
                    default:
                       //var_export($oNode->tagName);
                }
            }
        }
    }
    
    /**
     *
     * @param XMLReader $document
     * @param \DOMElement $node
     * @param string $baseFile
     */
    protected function loadShapeDrawing(XMLReader $document, \DOMElement $node, $baseFile)
    {
        // Core
        $oShape = new MemoryDrawing();
        $oShape->getShadow()->setVisible(false);
        // Variables
        $fileRels = 'ppt/slides/_rels/'.$baseFile.'.rels';
        
        $oElement = $document->getElement('p:nvPicPr/p:cNvPr', $node);
        if ($oElement) {
            $oShape->setName($oElement->hasAttribute('name') ? $oElement->getAttribute('name') : '');
            $oShape->setDescription($oElement->hasAttribute('descr') ? $oElement->getAttribute('descr') : '');
        }
        
        $oElement = $document->getElement('p:blipFill/a:blip', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('r:embed') && isset($this->arrayRels[$fileRels][$oElement->getAttribute('r:embed')])) {
                $pathImage = 'ppt/slides/'.$this->arrayRels[$fileRels][$oElement->getAttribute('r:embed')];
                $pathImage = explode('/', $pathImage);
                foreach ($pathImage as $key => $partPath) {
                    if ($partPath == '..') {
                        unset($pathImage[$key - 1]);
                        unset($pathImage[$key]);
                    }
                }
                $pathImage = implode('/', $pathImage);
                $imageFile = $this->oZip->getFromName($pathImage);
                if (!empty($imageFile)) {
                    $oShape->setImageResource(imagecreatefromstring($imageFile));
                }
            }
        }
        
        $oElement = $document->getElement('p:spPr/a:xfrm', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('rot')) {
                $oShape->setRotation(CommonDrawing::angleToDegrees($oElement->getAttribute('rot')));
            }
        }
        
        $oElement = $document->getElement('p:spPr/a:xfrm/a:off', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('x')) {
                $oShape->setOffsetX(CommonDrawing::emuToPixels($oElement->getAttribute('x')));
            }
            if ($oElement->hasAttribute('y')) {
                $oShape->setOffsetY(CommonDrawing::emuToPixels($oElement->getAttribute('y')));
            }
        }
        
        $oElement = $document->getElement('p:spPr/a:xfrm/a:ext', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('cx')) {
                $oShape->setWidth(CommonDrawing::emuToPixels($oElement->getAttribute('cx')));
            }
            if ($oElement->hasAttribute('cy')) {
                $oShape->setHeight(CommonDrawing::emuToPixels($oElement->getAttribute('cy')));
            }
        }
        
        $oElement = $document->getElement('p:spPr/a:effectLst', $node);
        if ($oElement) {
            $oShape->getShadow()->setVisible(true);
            
            $oSubElement = $document->getElement('a:outerShdw', $oElement);
            if ($oSubElement) {
                if ($oSubElement->hasAttribute('blurRad')) {
                    $oShape->getShadow()->setBlurRadius(CommonDrawing::emuToPixels($oSubElement->getAttribute('blurRad')));
                }
                if ($oSubElement->hasAttribute('dist')) {
                    $oShape->getShadow()->setDistance(CommonDrawing::emuToPixels($oSubElement->getAttribute('dist')));
                }
                if ($oSubElement->hasAttribute('dir')) {
                    $oShape->getShadow()->setDirection(CommonDrawing::angleToDegrees($oSubElement->getAttribute('dir')));
                }
                if ($oSubElement->hasAttribute('algn')) {
                    $oShape->getShadow()->setAlignment($oSubElement->getAttribute('algn'));
                }
            }
            
            $oSubElement = $document->getElement('a:outerShdw/a:srgbClr', $oElement);
            if ($oSubElement) {
                if ($oSubElement->hasAttribute('val')) {
                    $oColor = new Color();
                    $oColor->setRGB($oSubElement->getAttribute('val'));
                    $oShape->getShadow()->setColor($oColor);
                }
            }
            
            $oSubElement = $document->getElement('a:outerShdw/a:srgbClr/a:alpha', $oElement);
            if ($oSubElement) {
                if ($oSubElement->hasAttribute('val')) {
                    $oShape->getShadow()->setAlpha((int) $oSubElement->getAttribute('val') / 1000);
                }
            }
        }
        
        $this->oPhpPresentation->getActiveSlide()->addShape($oShape);
    }
    
    protected function loadShapeRichText(XMLReader $document, \DOMElement $node, $baseFile)
    {
        // Core
        $oShape = $this->oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->setParagraphs(array());
        // Variables
        $fileRels = 'ppt/slides/_rels/'.$baseFile.'.rels';
        
        $oElement = $document->getElement('p:spPr/a:xfrm', $node);
        if ($oElement && $oElement->hasAttribute('rot')) {
            $oShape->setRotation(CommonDrawing::angleToDegrees($oElement->getAttribute('rot')));
        }
        
        $oElement = $document->getElement('p:spPr/a:xfrm/a:off', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('x')) {
                $oShape->setOffsetX(CommonDrawing::emuToPixels($oElement->getAttribute('x')));
            }
            if ($oElement->hasAttribute('y')) {
                $oShape->setOffsetY(CommonDrawing::emuToPixels($oElement->getAttribute('y')));
            }
        }
        
        $oElement = $document->getElement('p:spPr/a:xfrm/a:ext', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('cx')) {
                $oShape->setWidth(CommonDrawing::emuToPixels($oElement->getAttribute('cx')));
            }
            if ($oElement->hasAttribute('cy')) {
                $oShape->setHeight(CommonDrawing::emuToPixels($oElement->getAttribute('cy')));
            }
        }
        
        $arrayElements = $document->getElements('p:txBody/a:p', $node);
        foreach ($arrayElements as $oElement) {
            // Core
            $oParagraph = $oShape->createParagraph();
            $oParagraph->setRichTextElements(array());
            
            $oSubElement = $document->getElement('a:pPr', $oElement);
            if ($oSubElement) {
                if ($oSubElement->hasAttribute('algn')) {
                    $oParagraph->getAlignment()->setHorizontal($oSubElement->getAttribute('algn'));
                }
                if ($oSubElement->hasAttribute('fontAlgn')) {
                    $oParagraph->getAlignment()->setVertical($oSubElement->getAttribute('fontAlgn'));
                }
                if ($oSubElement->hasAttribute('marL')) {
                    $oParagraph->getAlignment()->setMarginLeft(CommonDrawing::emuToPixels($oSubElement->getAttribute('marL')));
                }
                if ($oSubElement->hasAttribute('marR')) {
                    $oParagraph->getAlignment()->setMarginRight(CommonDrawing::emuToPixels($oSubElement->getAttribute('marR')));
                }
                if ($oSubElement->hasAttribute('indent')) {
                    $oParagraph->getAlignment()->setIndent(CommonDrawing::emuToPixels($oSubElement->getAttribute('indent')));
                }
                if ($oSubElement->hasAttribute('lvl')) {
                    $oParagraph->getAlignment()->setLevel($oSubElement->getAttribute('lvl'));
                }
                
                $oElementBuFont = $document->getElement('a:buFont', $oSubElement);
                $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_NONE);
                if ($oElementBuFont) {
                    if ($oElementBuFont->hasAttribute('typeface')) {
                        $oParagraph->getBulletStyle()->setBulletFont($oElementBuFont->getAttribute('typeface'));
                    }
                }
                $oElementBuChar = $document->getElement('a:buChar', $oSubElement);
                if ($oElementBuChar) {
                    $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
                    if ($oElementBuChar->hasAttribute('char')) {
                        $oParagraph->getBulletStyle()->setBulletChar($oElementBuChar->getAttribute('char'));
                    }
                }
                /*$oElementBuAutoNum = $document->getElement('a:buAutoNum', $oSubElement);
                if ($oElementBuAutoNum) {
                    $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_NUMERIC);
                    if ($oElementBuAutoNum->hasAttribute('type')) {
                        $oParagraph->getBulletStyle()->setBulletNumericStyle($oElementBuAutoNum->getAttribute('type'));
                    }
                    if ($oElementBuAutoNum->hasAttribute('startAt') && $oElementBuAutoNum->getAttribute('startAt') != 1) {
                        $oParagraph->getBulletStyle()->setBulletNumericStartAt($oElementBuAutoNum->getAttribute('startAt'));
                    }
                }*/
            }
            $arraySubElements = $document->getElements('(a:r|a:br)', $oElement);
            foreach ($arraySubElements as $oSubElement) {
                if ($oSubElement->tagName == 'a:br') {
                    $oParagraph->createBreak();
                }
                if ($oSubElement->tagName == 'a:r') {
                    $oElementrPr = $document->getElement('a:rPr', $oSubElement);
                    if (is_object($oElementrPr)) {
                        $oText = $oParagraph->createTextRun();

                        if ($oElementrPr->hasAttribute('b')) {
                            $oText->getFont()->setBold($oElementrPr->getAttribute('b') == 'true' ? true : false);
                        }
                        if ($oElementrPr->hasAttribute('i')) {
                            $oText->getFont()->setItalic($oElementrPr->getAttribute('i') == 'true' ? true : false);
                        }
                        if ($oElementrPr->hasAttribute('strike')) {
                            $oText->getFont()->setStrikethrough($oElementrPr->getAttribute('strike') == 'noStrike' ? false : true);
                        }
                        if ($oElementrPr->hasAttribute('sz')) {
                            $oText->getFont()->setSize((int)($oElementrPr->getAttribute('sz') / 100));
                        }
                        if ($oElementrPr->hasAttribute('u')) {
                            $oText->getFont()->setUnderline($oElementrPr->getAttribute('u'));
                        }
                        // Color
                        $oElementSrgbClr = $document->getElement('a:solidFill/a:srgbClr', $oElementrPr);
                        if (is_object($oElementSrgbClr) && $oElementSrgbClr->hasAttribute('val')) {
                            $oColor = new Color();
                            $oColor->setRGB($oElementSrgbClr->getAttribute('val'));
                            $oText->getFont()->setColor($oColor);
                        }
                        // Hyperlink
                        $oElementHlinkClick = $document->getElement('a:hlinkClick', $oElementrPr);
                        if (is_object($oElementHlinkClick)) {
                            if ($oElementHlinkClick->hasAttribute('tooltip')) {
                                $oText->getHyperlink()->setTooltip($oElementHlinkClick->getAttribute('tooltip'));
                            }
                            if ($oElementHlinkClick->hasAttribute('r:id') && isset($this->arrayRels[$fileRels][$oElementHlinkClick->getAttribute('r:id')])) {
                                $oText->getHyperlink()->setUrl($this->arrayRels[$fileRels][$oElementHlinkClick->getAttribute('r:id')]);
                            }
                        }
                    //} else {
                        // $oText = $oParagraph->createText();
                    }

                    $oSubSubElement = $document->getElement('a:t', $oSubElement);
                    $oText->setText($oSubSubElement->nodeValue);
                }
            }
        }
        
        if (count($oShape->getParagraphs()) > 0) {
            $oShape->setActiveParagraph(0);
        }
    }
    
    /**
     *
     * @param string $fileRels
     * @return string
     */
    protected function loadRels($fileRels)
    {
        $sPart = $this->oZip->getFromName($fileRels);
        if ($sPart !== false) {
            $xmlReader = new XMLReader();
            if ($xmlReader->getDomFromString($sPart)) {
                foreach ($xmlReader->getElements('*') as $oNode) {
                    $this->arrayRels[$fileRels][$oNode->getAttribute('Id')] = $oNode->getAttribute('Target');
                }
            }
        }
    }
}
