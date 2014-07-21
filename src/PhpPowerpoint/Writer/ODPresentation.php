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

use PhpOffice\PhpPowerpoint\HashTable;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Content;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Drawing;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Manifest;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Meta;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Mimetype;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\ObjectsChart;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Styles;
use PhpOffice\PhpPowerpoint\Shape\AbstractDrawing;

/**
 * ODPresentation writer
 */
class ODPresentation implements WriterInterface
{
    /**
    * Private PHPPowerPoint
    *
    * @var PHPPowerPoint
    */
    private $presentation;

    /**
    * Private writer parts
    *
    * @var \PhpOffice\PhpPowerpoint\Writer\ODPresentation\AbstractPart[]
    */
    private $writerParts;

    /**
     * Private unique hashtable
     *
     * @var \PhpOffice\PhpPowerpoint\HashTable
     */
    private $drawingHashTable;

    /**
     * @var \PhpOffice\PhpPowerpoint\Shape\Chart[]
     */
    public $chartArray = array();

    /**
    * Use disk caching where possible?
    *
    * @var boolean
    */
    private $useDiskCaching = false;

    /**
     * Disk caching directory
     *
     * @var string
     */
    private $diskCachingDirectory;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Writer\ODPresentation
     *
     * @param PHPPowerPoint $pPHPPowerPoint
     */
    public function __construct(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Assign PHPPowerPoint
        $this->setPHPPowerPoint($pPHPPowerPoint);

        // Set up disk caching location
        $this->diskCachingDirectory = './';

        // Initialise writer parts
        $this->writerParts['content']  = new Content();
        $this->writerParts['manifest'] = new Manifest();
        $this->writerParts['meta']     = new Meta();
        $this->writerParts['mimetype'] = new Mimetype();
        $this->writerParts['styles']   = new Styles();
        $this->writerParts['charts']   = new ObjectsChart();
        $this->writerParts['drawing']  = new Drawing();

        // Assign parent WriterInterface
        foreach ($this->writerParts as $writer) {
            $writer->setParentWriter($this);
        }

        // Set HashTable variables
        $this->drawingHashTable            = new HashTable();
    }

    /**
     * Save PHPPowerPoint to file
     *
     * @param  string    $pFilename
     * @throws \Exception
     */
    public function save($pFilename)
    {
        if (empty($pFilename)) {
            throw new \Exception("Filename is empty");
        }
        if (!is_null($this->presentation)) {
            // If $pFilename is php://output or php://stdout, make it a temporary file...
            $originalFilename = $pFilename;
            if (strtolower($pFilename) == 'php://output' || strtolower($pFilename) == 'php://stdout') {
                $pFilename = @tempnam('./', 'phppttmp');
                if ($pFilename == '') {
                    $pFilename = $originalFilename;
                }
            }

            $writerPartDrawing = $this->getWriterPart('Drawing');
            if (!$writerPartDrawing instanceof Drawing) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\ODPresentation\Drawing');
            }

            // Create drawing dictionary
            $this->drawingHashTable->addFromSource($writerPartDrawing->allDrawings($this->presentation));

            // Create new ZIP file and open it for writing
            $objZip = new \ZipArchive();

            // Try opening the ZIP file
            if ($objZip->open($pFilename, \ZIPARCHIVE::OVERWRITE) !== true) {
                if ($objZip->open($pFilename, \ZIPARCHIVE::CREATE) !== true) {
                    throw new \Exception("Could not open " . $pFilename . " for writing.");
                }
            }

            // Add mimetype to ZIP file
            //@todo Not in ZIPARCHIVE::CM_STORE mode
            $objZip->addFromString('mimetype', $this->getWriterPart('mimetype')->writePart());

            // Add content.xml to ZIP file
            $objZip->addFromString('content.xml', $this->getWriterPart('content')->writePart($this->presentation));

            // Add meta.xml to ZIP file
            $objZip->addFromString('meta.xml', $this->getWriterPart('meta')->writePart($this->presentation));

            // Add styles.xml to ZIP file
            $objZip->addFromString('styles.xml', $this->getWriterPart('styles')->writePart($this->presentation));

            // Add META-INF/manifest.xml
            $objZip->addFromString('META-INF/manifest.xml', $this->getWriterPart('manifest')->writePart());

            // Add charts
            foreach ($this->chartArray as $keyChart => $shapeChart) {
                $arrayFile = $this->getWriterPart('charts')->writePart($shapeChart);
                foreach ($arrayFile as $file => $content) {
                    if (!empty($content)) {
                        $objZip->addFromString('Object '.$keyChart.'/' . $file, $content);
                    }
                }
            }
            
            // Add media
            $arrMedia = array();
            for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
                $shape = $this->getDrawingHashTable()->getByIndex($i);
                if (!($shape instanceof AbstractDrawing)) {
                    throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Shape\AbstractDrawing');
                }
                if ($shape instanceof ShapeDrawing) {
                    if (!in_array(md5($shape->getPath()), $arrMedia)) {
                        $arrMedia[] = md5($shape->getPath());

                        $imagePath = $shape->getPath();

                        if (strpos($imagePath, 'zip://') !== false) {
                            $imagePath = substr($imagePath, 6);
                            $imagePathSplitted = explode('#', $imagePath);

                            $imageZip = new \ZipArchive();
                            $imageZip->open($imagePathSplitted[0]);
                            $imageContents = $imageZip->getFromName($imagePathSplitted[1]);
                            $imageZip->close();
                            unset($imageZip);
                        } else {
                            $imageContents = file_get_contents($imagePath);
                        }

                        $objZip->addFromString('Pictures/' . md5($shape->getPath()).'.'.$shape->getExtension(), $imageContents);
                    }
                } elseif ($shape instanceof MemoryDrawing) {
                    if (!in_array(str_replace(' ', '_', $shape->getIndexedFilename()), $arrMedia)) {
                        $arrMedia[] = str_replace(' ', '_', $shape->getIndexedFilename());
                        ob_start();
                            call_user_func($shape->getRenderingFunction(), $shape->getImageResource());
                            $imageContents = ob_get_contents();
                        ob_end_clean();

                        $objZip->addFromString('Pictures/' . str_replace(' ', '_', $shape->getIndexedFilename()), $imageContents);
                    }
                }
            }

            // Close file
            if ($objZip->close() === false) {
                throw new \Exception("Could not close zip file $pFilename.");
            }

            // If a temporary file was used, copy it to the correct file stream
            if ($originalFilename != $pFilename) {
                if (copy($pFilename, $originalFilename) === false) {
                    throw new \Exception("Could not copy temporary zip file $pFilename to $originalFilename.");
                }
                @unlink($pFilename);
            }
        } else {
            throw new \Exception("PHPPowerPoint object unassigned.");
        }
    }

    /**
     * Get PHPPowerPoint object
     *
     * @return PHPPowerPoint
     * @throws \Exception
     */
    public function getPHPPowerPoint()
    {
        if (!is_null($this->presentation)) {
            return $this->presentation;
        } else {
            throw new \Exception("No PHPPowerPoint assigned.");
        }
    }

    /**
     * Get PHPPowerPoint object
     *
     * @param  PHPPowerPoint                       $pPHPPowerPoint PHPPowerPoint object
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007
     */
    public function setPHPPowerPoint(PHPPowerPoint $pPHPPowerPoint = null)
    {
        $this->presentation = $pPHPPowerPoint;

        return $this;
    }

    /**
     * Get drawing hash table
     *
     * @return \PhpOffice\PhpPowerpoint\HashTable
     */
    public function getDrawingHashTable()
    {
        return $this->drawingHashTable;
    }

    /**
     * Get writer part
     *
     * @param  string                                         $pPartName Writer part name
     * @return \PhpOffice\PhpPowerpoint\Writer\ODPresentation\AbstractPart
     */
    public function getWriterPart($pPartName = '')
    {
        if ($pPartName != '' && isset($this->writerParts[strtolower($pPartName)])) {
            return $this->writerParts[strtolower($pPartName)];
        } else {
            return null;
        }
    }

    /**
     * Get use disk caching where possible?
     *
     * @return boolean
     */
    public function hasDiskCaching()
    {
        return $this->useDiskCaching;
    }

    /**
     * Set use disk caching where possible?
     *
     * @param  boolean $pValue
     * @param  string $pDirectory Disk caching directory
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007
     */
    public function setUseDiskCaching($pValue = false, $pDirectory = null)
    {
        $this->useDiskCaching = $pValue;

        if (!is_null($pDirectory)) {
            if (is_dir($pDirectory)) {
                $this->diskCachingDirectory = $pDirectory;
            } else {
                throw new \Exception("Directory does not exist: $pDirectory");
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
        return $this->diskCachingDirectory;
    }
}
