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

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\HashTable;
use PhpOffice\PhpPowerpoint\AbstractShape;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Chart;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\ContentTypes;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\DocProps;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Drawing;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\AbstractLayoutPack;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\LayoutPack\PackDefault;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Presentation;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Rels;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Slide;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Theme;

/**
 * PHPPowerPoint_Writer_PowerPoint2007
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PowerPoint2007 implements WriterInterface
{
    /**
     * Office2003 compatibility
     *
     * @var boolean
     */
    protected $office2003comp = false;

    /**
     * Private writer parts
     *
     * @var PHPPowerPoint_Writer_PowerPoint2007_AbstractPart[]
     */
    protected $writerParts;

    /**
     * Private PHPPowerPoint
     *
     * @var PHPPowerPoint
     */
    protected $presentation;

    /**
     * Private unique PHPPowerPoint_Worksheet_BaseDrawing HashTable
     *
     * @var PHPPowerPoint_HashTable
     */
    protected $drawingHashTable;

    /**
     * Use disk caching where possible?
     *
     * @var boolean
     */
    protected $useDiskCaching = false;

    /**
     * Disk caching directory
     *
     * @var string
     */
    protected $diskCachingDir;

    /**
     * Layout pack to use
     *
     * @var PHPPowerPoint_Writer_PowerPoint2007_LayoutPack
     */
    protected $layoutPack;

    /**
     * Create a new PHPPowerPoint_Writer_PowerPoint2007
     *
     * @param PHPPowerPoint $pPHPPowerPoint
     */
    public function __construct(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Assign PHPPowerPoint
        $this->setPHPPowerPoint($pPHPPowerPoint);

        // Set up disk caching location
        $this->diskCachingDir = './';

        // Set layout pack
        $this->layoutPack = new PackDefault();

        // Initialise writer parts
        $this->writerParts['contenttypes'] = new ContentTypes();
        $this->writerParts['docprops']     = new DocProps();
        $this->writerParts['rels']         = new Rels();
        $this->writerParts['theme']        = new Theme();
        $this->writerParts['presentation'] = new Presentation();
        $this->writerParts['slide']        = new Slide();
        $this->writerParts['drawing']      = new Drawing();
        $this->writerParts['chart']        = new Chart();

        // Assign parent WriterInterface
        foreach ($this->writerParts as $writer) {
            $writer->setParentWriter($this);
        }

        // Set HashTable variables
        $this->drawingHashTable = new HashTable();
    }

    /**
     * Get writer part
     *
     * @param  string                                         $pPartName Writer part name
     * @return PHPPowerPoint_Writer_PowerPoint2007_AbstractPart
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

            // Create drawing dictionary
            $this->drawingHashTable->addFromSource($this->getWriterPart('Drawing')->allDrawings($this->presentation));

            // Create new ZIP file and open it for writing
            $objZip = new \ZipArchive();

            // Try opening the ZIP file
            if ($objZip->open($pFilename, \ZIPARCHIVE::OVERWRITE) !== true) {
                if ($objZip->open($pFilename, ZIPARCHIVE::CREATE) !== true) {
                    throw new \Exception("Could not open " . $pFilename . " for writing.");
                }
            }

            // Add [Content_Types].xml to ZIP file
            $objZip->addFromString('[Content_Types].xml', $this->getWriterPart('ContentTypes')->writeContentTypes($this->presentation));

            // Add relationships to ZIP file
            $objZip->addFromString('_rels/.rels', $this->getWriterPart('Rels')->writeRelationships());
            $objZip->addFromString('ppt/_rels/presentation.xml.rels', $this->getWriterPart('Rels')->writePresentationRelationships($this->presentation));

            // Add document properties to ZIP file
            $objZip->addFromString('docProps/app.xml', $this->getWriterPart('DocProps')->writeDocPropsApp($this->presentation));
            $objZip->addFromString('docProps/core.xml', $this->getWriterPart('DocProps')->writeDocPropsCore($this->presentation));

            // Add themes to ZIP file
            $masterSlides = $this->getLayoutPack()->getMasterSlides();
            foreach ($masterSlides as $masterSlide) {
                $objZip->addFromString('ppt/theme/_rels/theme' . $masterSlide['masterid'] . '.xml.rels', $this->getWriterPart('Rels')->writeThemeRelationships($masterSlide['masterid']));
                $objZip->addFromString('ppt/theme/theme' . $masterSlide['masterid'] . '.xml', utf8_encode($this->getWriterPart('Theme')->writeTheme($masterSlide['masterid'])));
            }

            // Add slide masters to ZIP file
            $masterSlides = $this->getLayoutPack()->getMasterSlides();
            foreach ($masterSlides as $masterSlide) {
                $objZip->addFromString('ppt/slideMasters/_rels/slideMaster' . $masterSlide['masterid'] . '.xml.rels', $this->getWriterPart('Rels')->writeSlideMasterRelationships($masterSlide['masterid']));
                $objZip->addFromString('ppt/slideMasters/slideMaster' . $masterSlide['masterid'] . '.xml', $masterSlide['body']);
            }

            // Add slide layouts to ZIP file
            $slideLayouts = $this->getLayoutPack()->getLayouts();
            foreach ($slideLayouts as $key => $layout) {
                $objZip->addFromString('ppt/slideLayouts/_rels/slideLayout' . $key . '.xml.rels', $this->getWriterPart('Rels')->writeSlideLayoutRelationships($key, $layout['masterid']));
                $objZip->addFromString('ppt/slideLayouts/slideLayout' . $key . '.xml', utf8_encode($layout['body']));
            }

            // Add layoutpack relations
            $otherRelations = $this->getLayoutPack()->getMasterSlideRelations();
            foreach ($otherRelations as $otherRelations) {
                if (strpos($otherRelations['target'], 'http://') !== 0) {
                    $objZip->addFromString($this->absoluteZipPath('ppt/slideMasters/' . $otherRelations['target']), $otherRelations['contents']);
                }
            }
            $otherRelations = $this->getLayoutPack()->getThemeRelations();
            foreach ($otherRelations as $otherRelations) {
                if (strpos($otherRelations['target'], 'http://') !== 0) {
                    $objZip->addFromString($this->absoluteZipPath('ppt/theme/' . $otherRelations['target']), $otherRelations['contents']);
                }
            }
            $otherRelations = $this->getLayoutPack()->getLayoutRelations();
            foreach ($otherRelations as $otherRelations) {
                if (strpos($otherRelations['target'], 'http://') !== 0) {
                    $objZip->addFromString($this->absoluteZipPath('ppt/slideLayouts/' . $otherRelations['target']), $otherRelations['contents']);
                }
            }

            // Add presentation to ZIP file
            $objZip->addFromString('ppt/presentation.xml', $this->getWriterPart('Presentation')->writePresentation($this->presentation));

            // Add slides (drawings, ...) and slide relationships (drawings, ...)
            for ($i = 0; $i < $this->presentation->getSlideCount(); ++$i) {
                // Add slide
                $objZip->addFromString('ppt/slides/_rels/slide' . ($i + 1) . '.xml.rels', $this->getWriterPart('Rels')->writeSlideRelationships($this->presentation->getSlide($i)));
                $objZip->addFromString('ppt/slides/slide' . ($i + 1) . '.xml', $this->getWriterPart('Slide')->writeSlide($this->presentation->getSlide($i)));
            }

            // Add media
            for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
                if ($this->getDrawingHashTable()->getByIndex($i) instanceof Shape\Drawing) {
                    $imagePath     = $this->getDrawingHashTable()->getByIndex($i)->getPath();

                    if (strpos($imagePath, 'zip://') !== false) {
                        $imagePath         = substr($imagePath, 6);
                        $imagePathSplitted = explode('#', $imagePath);

                        $imageZip = new \ZipArchive();
                        $imageZip->open($imagePathSplitted[0]);
                        $imageContents = $imageZip->getFromName($imagePathSplitted[1]);
                        $imageZip->close();
                        unset($imageZip);
                    } else {
                        $imageContents = file_get_contents($imagePath);
                    }

                    $objZip->addFromString('ppt/media/' . str_replace(' ', '_', $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename()), $imageContents);
                } elseif ($this->getDrawingHashTable()->getByIndex($i) instanceof Shape\MemoryDrawing) {
                    ob_start();
                    call_user_func($this->getDrawingHashTable()->getByIndex($i)->getRenderingFunction(), $this->getDrawingHashTable()->getByIndex($i)->getImageResource());
                    $imageContents = ob_get_contents();
                    ob_end_clean();

                    $objZip->addFromString('ppt/media/' . str_replace(' ', '_', $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename()), $imageContents);
                } elseif ($this->getDrawingHashTable()->getByIndex($i) instanceof Shape\Chart) {
                    $objZip->addFromString('ppt/charts/' . $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename(), $this->getWriterPart('Chart')->writeChart($this->getDrawingHashTable()->getByIndex($i)));

                    // Chart relations?
                    if ($this->getDrawingHashTable()->getByIndex($i)->hasIncludedSpreadsheet()) {
                        $objZip->addFromString('ppt/charts/_rels/' . $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename() . '.rels', $this->getWriterPart('Rels')->writeChartRelationships($this->getDrawingHashTable()->getByIndex($i)));
                        $objZip->addFromString('ppt/embeddings/' . $this->getDrawingHashTable()->getByIndex($i)->getIndexedFilename() . '.xlsx', $this->getWriterPart('Chart')->writeSpreadsheet($this->presentation, $this->getDrawingHashTable()->getByIndex($i), $pFilename . '.xlsx'));
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
     * @return PHPPowerPoint_Writer_PowerPoint2007
     */
    public function setPHPPowerPoint(PHPPowerPoint $pPHPPowerPoint = null)
    {
        $this->presentation = $pPHPPowerPoint;

        return $this;
    }

    /**
     * Get PHPPowerPoint_Worksheet_BaseDrawing HashTable
     *
     * @return PHPPowerPoint_HashTable
     */
    public function getDrawingHashTable()
    {
        return $this->drawingHashTable;
    }

    /**
     * Get Office2003 compatibility
     *
     * @return boolean
     */
    public function hasOffice2003Compatibility()
    {
        return $this->office2003comp;
    }

    /**
     * Set Pre-Calculate Formulas
     *
     * @param  boolean                             $pValue Office2003 compatibility?
     * @return PHPPowerPoint_Writer_PowerPoint2007
     */
    public function setOffice2003Compatibility($pValue = false)
    {
        $this->office2003comp = $pValue;

        return $this;
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
     * @param  boolean                             $pValue
     * @param  string                              $pDirectory Disk caching directory
     * @throws \Exception                           \Exception when directory does not exist
     * @return PHPPowerPoint_Writer_PowerPoint2007
     */
    public function setUseDiskCaching($pValue = false, $pDirectory = null)
    {
        $this->useDiskCaching = $pValue;

        if (!is_null($pDirectory)) {
            if (is_dir($pDirectory)) {
                $this->diskCachingDir = $pDirectory;
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
        return $this->diskCachingDir;
    }

    /**
     * Get layout pack to use
     *
     * @return PHPPowerPoint_Writer_PowerPoint2007_LayoutPack
     */
    public function getLayoutPack()
    {
        return $this->layoutPack;
    }

    /**
     * Set layout pack to use
     *
     * @param  PHPPowerPoint_Writer_PowerPoint2007_LayoutPack $pValue
     * @return PHPPowerPoint_Writer_PowerPoint2007
     */
    public function setLayoutPack(AbstractLayoutPack $pValue = null)
    {
        $this->layoutPack = $pValue;

        return $this;
    }

    /**
     * Determine absolute zip path
     *
     * @param  string $path
     * @return string
     */
    protected function absoluteZipPath($path)
    {
        $path      = str_replace(array(
            '/',
            '\\'
        ), DIRECTORY_SEPARATOR, $path);
        $parts     = array_filter(explode(DIRECTORY_SEPARATOR, $path), 'strlen');
        $absolutes = array();
        foreach ($parts as $part) {
            if ('.' == $part) {
                continue;
            }
            if ('..' == $part) {
                array_pop($absolutes);
            } else {
                $absolutes[] = $part;
            }
        }

        return implode('/', $absolutes);
    }
}
