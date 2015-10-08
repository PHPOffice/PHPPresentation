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

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
use PhpOffice\PhpPresentation\Writer\ODPresentation\Content;
use PhpOffice\PhpPresentation\Writer\ODPresentation\Drawing;
use PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest;
use PhpOffice\PhpPresentation\Writer\ODPresentation\Meta;
use PhpOffice\PhpPresentation\Writer\ODPresentation\Mimetype;
use PhpOffice\PhpPresentation\Writer\ODPresentation\ObjectsChart;
use PhpOffice\PhpPresentation\Writer\ODPresentation\Styles;
use PhpOffice\PhpPresentation\Shape\AbstractDrawing;

/**
 * ODPresentation writer
 */
class ODPresentation implements WriterInterface
{
    /**
    * Private PhpPresentation
    *
    * @var \PhpOffice\PhpPresentation\PhpPresentation
    */
    private $presentation;

    /**
    * Private writer parts
    *
    * @var \PhpOffice\PhpPresentation\Writer\ODPresentation\AbstractPart[]
    */
    private $writerParts;

    /**
     * Private unique hashtable
     *
     * @var \PhpOffice\PhpPresentation\HashTable
     */
    private $drawingHashTable;

    /**
     * @var \PhpOffice\PhpPresentation\Shape\Chart[]
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
     * Create a new \PhpOffice\PhpPresentation\Writer\ODPresentation
     *
     * @param PhpPresentation $pPhpPresentation
     */
    public function __construct(PhpPresentation $pPhpPresentation = null)
    {
        // Assign PhpPresentation
        $this->setPhpPresentation($pPhpPresentation);

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
     * Save PhpPresentation to file
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

            $writerPartChart = $this->getWriterPart('charts');
            if (!$writerPartChart instanceof ObjectsChart) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\ODPresentation\ObjectsChart');
            }
            $writerPartContent = $this->getWriterPart('content');
            if (!$writerPartContent instanceof Content) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\ODPresentation\Content');
            }
            $writerPartDrawing = $this->getWriterPart('Drawing');
            if (!$writerPartDrawing instanceof Drawing) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\ODPresentation\Drawing');
            }
            $writerPartManifest = $this->getWriterPart('manifest');
            if (!$writerPartManifest instanceof Manifest) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest');
            }
            $writerPartMeta = $this->getWriterPart('meta');
            if (!$writerPartMeta instanceof Meta) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\ODPresentation\Meta');
            }
            $writerPartMimetype = $this->getWriterPart('mimetype');
            if (!$writerPartMimetype instanceof Mimetype) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\ODPresentation\Mimetype');
            }
            $writerPartStyles = $this->getWriterPart('styles');
            if (!$writerPartStyles instanceof Styles) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\ODPresentation\Styles');
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
            $objZip->addFromString('mimetype', $writerPartMimetype->writePart());

            // Add content.xml to ZIP file
            $objZip->addFromString('content.xml', $writerPartContent->writePart($this->presentation));

            // Add meta.xml to ZIP file
            $objZip->addFromString('meta.xml', $writerPartMeta->writePart($this->presentation));

            // Add styles.xml to ZIP file
            $objZip->addFromString('styles.xml', $writerPartStyles->writePart($this->presentation));

            // Add META-INF/manifest.xml
            $objZip->addFromString('META-INF/manifest.xml', $writerPartManifest->writePart());

            // Add charts
            foreach ($this->chartArray as $keyChart => $shapeChart) {
                $arrayFile = $writerPartChart->writePart($shapeChart);
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
                    throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Shape\AbstractDrawing');
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
                if (@unlink($pFilename) === false) {
                    throw new \Exception('The file '.$pFilename.' could not be deleted.');
                }
            }
        } else {
            throw new \Exception("PhpPresentation object unassigned.");
        }
    }

    /**
     * Get PhpPresentation object
     *
     * @return PhpPresentation
     * @throws \Exception
     */
    public function getPhpPresentation()
    {
        if (!is_null($this->presentation)) {
            return $this->presentation;
        } else {
            throw new \Exception("No PhpPresentation assigned.");
        }
    }

    /**
     * Get PhpPresentation object
     *
     * @param  PhpPresentation                       $pPhpPresentation PhpPresentation object
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Writer\ODPresentation
     */
    public function setPhpPresentation(PhpPresentation $pPhpPresentation = null)
    {
        $this->presentation = $pPhpPresentation;

        return $this;
    }

    /**
     * Get drawing hash table
     *
     * @return \PhpOffice\PhpPresentation\HashTable
     */
    public function getDrawingHashTable()
    {
        return $this->drawingHashTable;
    }

    /**
     * Get writer part
     *
     * @param  string                                         $pPartName Writer part name
     * @return \PhpOffice\PhpPresentation\Writer\ODPresentation\AbstractPart
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
     * @return \PhpOffice\PhpPresentation\Writer\ODPresentation
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
