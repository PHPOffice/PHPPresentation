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
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */

/**
 * ODPresentation writer
 */
class PHPPowerPoint_Writer_ODPresentation implements PHPPowerPoint_Writer_IWriter
{
    /**
    * Private PHPPowerPoint
    *
    * @var PHPPowerPoint
    */
    private $_presentation;

    /**
    * Private writer parts
    *
    * @var PHPPowerPoint_Writer_ODPresentation_WriterPart[]
    */
    private $_writerParts;

    /**
     * Private unique PHPPowerPoint_Worksheet_BaseDrawing HashTable
     *
     * @var PHPPowerPoint_HashTable
     */
    private $_drawingHashTable;

    /**
    * Use disk caching where possible?
    *
    * @var boolean
    */
    private $_useDiskCaching = false;

    /**
     * Disk caching directory
     *
     * @var string
     */
    private $_diskCachingDirectory;

    /**
     * Create a new PHPPowerPoint_Writer_ODPresentation
     *
     * @param PHPPowerPoint $pPHPPowerPoint
     */
    public function __construct(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Assign PHPPowerPoint
        $this->setPHPPowerPoint($pPHPPowerPoint);

        // Set up disk caching location
        $this->_diskCachingDirectory = './';

        // Initialise writer parts
        $this->_writerParts['content']  = new PHPPowerPoint_Writer_ODPresentation_Content();
        $this->_writerParts['manifest'] = new PHPPowerPoint_Writer_ODPresentation_Manifest();
        $this->_writerParts['meta']     = new PHPPowerPoint_Writer_ODPresentation_Meta();
        $this->_writerParts['mimetype'] = new PHPPowerPoint_Writer_ODPresentation_Mimetype();
        $this->_writerParts['styles']   = new PHPPowerPoint_Writer_ODPresentation_Styles();

        $this->_writerParts['drawing']  = new PHPPowerPoint_Writer_ODPresentation_Drawing();

        // Assign parent IWriter
        foreach ($this->_writerParts as $writer) {
            $writer->setParentWriter($this);
        }

        // Set HashTable variables
        $this->_drawingHashTable            = new PHPPowerPoint_HashTable();
    }

    /**
     * Save PHPPowerPoint to file
     *
     * @param  string    $pFilename
     * @throws Exception
     */
    public function save($pFilename = null)
    {
        if (!is_null($this->_presentation)) {
            // If $pFilename is php://output or php://stdout, make it a temporary file...
            $originalFilename = $pFilename;
            if (strtolower($pFilename) == 'php://output' || strtolower($pFilename) == 'php://stdout') {
                $pFilename = @tempnam('./', 'phppttmp');
                if ($pFilename == '') {
                    $pFilename = $originalFilename;
                }
            }

            // Create drawing dictionary
            $this->_drawingHashTable->addFromSource($this->getWriterPart('Drawing')->allDrawings($this->_presentation));

            // Create new ZIP file and open it for writing
            $objZip = new ZipArchive();

            // Try opening the ZIP file
            if ($objZip->open($pFilename, ZIPARCHIVE::OVERWRITE) !== true) {
                if ($objZip->open($pFilename, ZIPARCHIVE::CREATE) !== true) {
                    throw new Exception("Could not open " . $pFilename . " for writing.");
                }
            }

            // Add mimetype to ZIP file
            //@todo Not in ZIPARCHIVE::CM_STORE mode
            $objZip->addFromString('mimetype', $this->getWriterPart('mimetype')->writeMimetype());

            // Add content.xml to ZIP file
            $objZip->addFromString('content.xml', $this->getWriterPart('content')->writeContent($this->_presentation));

            // Add meta.xml to ZIP file
            $objZip->addFromString('meta.xml', $this->getWriterPart('meta')->writeMeta($this->_presentation));

            // Add styles.xml to ZIP file
            $objZip->addFromString('styles.xml', $this->getWriterPart('styles')->writeStyles($this->_presentation));

            // Add META-INF/manifest.xml
            $objZip->addFromString('META-INF/manifest.xml', $this->getWriterPart('manifest')->writeManifest());

            // Add media
            $arrMedia = array();
            for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
                if ($this->getDrawingHashTable()->getByIndex($i) instanceof PHPPowerPoint_Shape_Drawing) {
                    if (!in_array(md5($this->getDrawingHashTable()->getByIndex($i)->getPath()), $arrMedia)) {
                        $arrMedia[] = md5($this->getDrawingHashTable()->getByIndex($i)->getPath());

                        $imageContents = null;
                        $imagePath = $this->getDrawingHashTable()->getByIndex($i)->getPath();

                        if (strpos($imagePath, 'zip://') !== false) {
                            $imagePath = substr($imagePath, 6);
                            $imagePathSplitted = explode('#', $imagePath);

                            $imageZip = new ZipArchive();
                            $imageZip->open($imagePathSplitted[0]);
                            $imageContents = $imageZip->getFromName($imagePathSplitted[1]);
                            $imageZip->close();
                            unset($imageZip);
                        } else {
                            $imageContents = file_get_contents($imagePath);
                        }

                        $objZip->addFromString('Pictures/' . md5($this->getDrawingHashTable()->getByIndex($i)->getPath()).'.'.$this->getDrawingHashTable()->getByIndex($i)->getExtension(), $imageContents);
                    }
                } elseif ($this->getDrawingHashTable()->getByIndex($i) instanceof PHPPowerPoint_Shape_MemoryDrawing) {
                    if (!in_array(md5($this->getDrawingHashTable()->getByIndex($i)->getPath()), $arrMedia)) {
                        $arrMedia[] = md5($this->getDrawingHashTable()->getByIndex($i)->getPath());
                        ob_start();
                            call_user_func($this->getDrawingHashTable()->getByIndex($i)->getRenderingFunction(), $this->getDrawingHashTable()->getByIndex($i)->getImageResource());
                            $imageContents = ob_get_contents();
                        ob_end_clean();

                        $objZip->addFromString('Pictures/' . md5($this->getDrawingHashTable()->getByIndex($i)->getPath()).'.'.$this->getDrawingHashTable()->getByIndex($i)->getExtension(), $imageContents);
                    }
                }
            }

            // Close file
            if ($objZip->close() === false) {
                throw new Exception("Could not close zip file $pFilename.");
            }

            // If a temporary file was used, copy it to the correct file stream
            if ($originalFilename != $pFilename) {
                if (copy($pFilename, $originalFilename) === false) {
                    throw new Exception("Could not copy temporary zip file $pFilename to $originalFilename.");
                }
                @unlink($pFilename);
            }

        } else {
            throw new Exception("PHPPowerPoint object unassigned.");
        }
    }

    /**
     * Get PHPPowerPoint object
     *
     * @return PHPPowerPoint
     * @throws Exception
     */
    public function getPHPPowerPoint()
    {
        if (!is_null($this->_presentation)) {
            return $this->_presentation;
        } else {
            throw new Exception("No PHPPowerPoint assigned.");
        }
    }

    /**
     * Get PHPPowerPoint object
     *
     * @param  PHPPowerPoint                       $pPHPPowerPoint PHPPowerPoint object
     * @throws Exception
     * @return PHPPowerPoint_Writer_PowerPoint2007
     */
    public function setPHPPowerPoint(PHPPowerPoint $pPHPPowerPoint = null)
    {
        $this->_presentation = $pPHPPowerPoint;

        return $this;
    }

    /**
     * Get PHPPowerPoint_Worksheet_BaseDrawing HashTable
     *
     * @return PHPPowerPoint_HashTable
     */
    public function getDrawingHashTable()
    {
        return $this->_drawingHashTable;
    }

    /**
     * Get writer part
     *
     * @param  string                                         $pPartName Writer part name
     * @return PHPPowerPoint_Writer_ODPresentation_WriterPart
     */
    public function getWriterPart($pPartName = '')
    {
        if ($pPartName != '' && isset($this->_writerParts[strtolower($pPartName)])) {
            return $this->_writerParts[strtolower($pPartName)];
        } else {
            return null;
        }
    }

    /**
     * Get use disk caching where possible?
     *
     * @return boolean
     */
    public function getUseDiskCaching()
    {
        return $this->_useDiskCaching;
    }

    /**
     * Set use disk caching where possible?
     *
     * @param  boolean                             $pValue
     * @param  string                              $pDirectory Disk caching directory
     * @throws Exception                           Exception when directory does not exist
     * @return PHPPowerPoint_Writer_PowerPoint2007
     */
    public function setUseDiskCaching($pValue = false, $pDirectory = null)
    {
        $this->_useDiskCaching = $pValue;

        if (!is_null($pDirectory)) {
            if (is_dir($pDirectory)) {
                $this->_diskCachingDirectory = $pDirectory;
            } else {
                throw new Exception("Directory does not exist: $pDirectory");
            }
        }

        return $this;
    }

    /**
     * Get disk caching directory
     *
     * @return string
     */
    public function getDiskCachingDirectory()
    {
        return $this->_diskCachingDirectory;
    }
}
