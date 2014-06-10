<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPowerpoint\Writer;

use PhpOffice\PhpPowerpoint\Writer\IWriter;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Content;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Manifest;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Meta;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Mimetype;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Styles;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Drawing;
use PhpOffice\PhpPowerpoint\HashTable;
use PhpOffice\PhpPowerpoint\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;

/**
 * ODPresentation writer
 */
class ODPresentation implements IWriter
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
        $this->_writerParts['content']  = new Content();
        $this->_writerParts['manifest'] = new Manifest();
        $this->_writerParts['meta']     = new Meta();
        $this->_writerParts['mimetype'] = new Mimetype();
        $this->_writerParts['styles']   = new Styles();

        $this->_writerParts['drawing']  = new Drawing();

        // Assign parent IWriter
        foreach ($this->_writerParts as $writer) {
            $writer->setParentWriter($this);
        }

        // Set HashTable variables
        $this->_drawingHashTable            = new HashTable();
    }

    /**
     * Save PHPPowerPoint to file
     *
     * @param  string    $pFilename
     * @throws Exception
     */
    public function save($pFilename)
    {
    	if (empty($pFilename)) {
    		throw new Exception("Filename is empty");
    	}
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
            $objZip = new \ZipArchive();

            // Try opening the ZIP file
            if ($objZip->open($pFilename, \ZIPARCHIVE::OVERWRITE) !== true) {
                if ($objZip->open($pFilename, \ZIPARCHIVE::CREATE) !== true) {
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
                if ($this->getDrawingHashTable()->getByIndex($i) instanceof ShapeDrawing) {
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
                } elseif ($this->getDrawingHashTable()->getByIndex($i) instanceof MemoryDrawing) {
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
