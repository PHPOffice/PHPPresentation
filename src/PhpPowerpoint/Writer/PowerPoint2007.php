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
use PhpOffice\PhpPowerpoint\Shape\Drawing as DrawingShape;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing as MemoryDrawingShape;
use PhpOffice\PhpPowerpoint\Shape\Chart as ChartShape;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\AbstractLayoutPack;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Chart;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\ContentTypes;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\DocProps;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Drawing;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\LayoutPack\PackDefault;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\PptProps;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Presentation;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Rels;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Slide;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Theme;

/**
 * \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007
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
     * @var \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\AbstractPart[]
     */
    protected $writerParts;

    /**
     * Private PHPPowerPoint
     *
     * @var PHPPowerPoint
     */
    protected $presentation;

    /**
     * Private unique hash table
     *
     * @var \PhpOffice\PhpPowerpoint\HashTable
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
     * @var \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\AbstractLayoutPack
     */
    protected $layoutPack;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007
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
        $this->writerParts['pptprops']     = new PptProps();
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
     * @return \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\AbstractPart
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
            
            $wPartDrawing = $this->getWriterPart('Drawing');
            if (!$wPartDrawing instanceof Drawing) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Drawing');
            }
            $wPartContentTypes = $this->getWriterPart('ContentTypes');
            if (!$wPartContentTypes instanceof ContentTypes) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\ContentTypes');
            }
            $wPartRels = $this->getWriterPart('Rels');
            if (!$wPartRels instanceof Rels) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Rels');
            }
            $wPartDocProps = $this->getWriterPart('DocProps');
            if (!$wPartDocProps instanceof DocProps) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\DocProps');
            }
            $wPartTheme = $this->getWriterPart('Theme');
            if (!$wPartTheme instanceof Theme) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Theme');
            }
            $wPartPresentation = $this->getWriterPart('Presentation');
            if (!$wPartPresentation instanceof Presentation) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Presentation');
            }
            $wPartSlide = $this->getWriterPart('Slide');
            if (!$wPartSlide instanceof Slide) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Slide');
            }
            $wPartChart = $this->getWriterPart('Chart');
            if (!$wPartChart instanceof Chart) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Chart');
            }
            $wPptProps = $this->getWriterPart('PptProps');
            if (!$wPptProps instanceof PptProps) {
                throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\PptProps');
            }
            
            // Create drawing dictionary
            $this->drawingHashTable->addFromSource($wPartDrawing->allDrawings($this->presentation));

            // Create new ZIP file and open it for writing
            $objZip = new \ZipArchive();

            // Try opening the ZIP file
            if ($objZip->open($pFilename, \ZipArchive::OVERWRITE) !== true) {
                if ($objZip->open($pFilename, \ZipArchive::CREATE) !== true) {
                    throw new \Exception("Could not open " . $pFilename . " for writing.");
                }
            }

            // Add [Content_Types].xml to ZIP file
            $objZip->addFromString('[Content_Types].xml', $wPartContentTypes->writeContentTypes($this->presentation));

            // Add PPT properties and styles to ZIP file - Required for Apple Keynote compatibility.
            $objZip->addFromString('ppt/presProps.xml', $wPptProps->writePresProps());
            $objZip->addFromString('ppt/tableStyles.xml', $wPptProps->writeTableStyles());
            $objZip->addFromString('ppt/viewProps.xml', $wPptProps->writeViewProps());

            // Add relationships to ZIP file
            $objZip->addFromString('_rels/.rels', $wPartRels->writeRelationships());
            $objZip->addFromString('ppt/_rels/presentation.xml.rels', $wPartRels->writePresentationRelationships($this->presentation));

            // Add document properties to ZIP file
            $objZip->addFromString('docProps/app.xml', $wPartDocProps->writeDocPropsApp($this->presentation));
            $objZip->addFromString('docProps/core.xml', $wPartDocProps->writeDocPropsCore($this->presentation));

            $masterSlides = $this->getLayoutPack()->getMasterSlides();
            foreach ($masterSlides as $masterSlide) {
                // Add themes to ZIP file
                $objZip->addFromString('ppt/theme/_rels/theme' . $masterSlide['masterid'] . '.xml.rels', $wPartRels->writeThemeRelationships($masterSlide['masterid']));
                $objZip->addFromString('ppt/theme/theme' . $masterSlide['masterid'] . '.xml', utf8_encode($wPartTheme->writeTheme($masterSlide['masterid'])));
                // Add slide masters to ZIP file
                $objZip->addFromString('ppt/slideMasters/_rels/slideMaster' . $masterSlide['masterid'] . '.xml.rels', $wPartRels->writeSlideMasterRelationships($masterSlide['masterid']));
                $objZip->addFromString('ppt/slideMasters/slideMaster' . $masterSlide['masterid'] . '.xml', $masterSlide['body']);
            }

            // Add slide layouts to ZIP file
            $slideLayouts = $this->getLayoutPack()->getLayouts();
            foreach ($slideLayouts as $key => $layout) {
                $objZip->addFromString('ppt/slideLayouts/_rels/slideLayout' . $key . '.xml.rels', $wPartRels->writeSlideLayoutRelationships($key, $layout['masterid']));
                $objZip->addFromString('ppt/slideLayouts/slideLayout' . $key . '.xml', utf8_encode($layout['body']));
            }

            // Add layoutpack relations
            $otherRelations = $this->getLayoutPack()->getMasterSlideRelations();
            foreach ($otherRelations as $otherRelation) {
                if (strpos($otherRelation['target'], 'http://') !== 0) {
                    $objZip->addFromString($this->absoluteZipPath('ppt/slideMasters/' . $otherRelation['target']), $otherRelation['contents']);
                }
            }
            $otherRelations = $this->getLayoutPack()->getThemeRelations();
            foreach ($otherRelations as $otherRelation) {
                if (strpos($otherRelation['target'], 'http://') !== 0) {
                    $objZip->addFromString($this->absoluteZipPath('ppt/theme/' . $otherRelation['target']), $otherRelation['contents']);
                }
            }
            $otherRelations = $this->getLayoutPack()->getLayoutRelations();
            foreach ($otherRelations as $otherRelation) {
                if (strpos($otherRelation['target'], 'http://') !== 0) {
                    $objZip->addFromString($this->absoluteZipPath('ppt/slideLayouts/' . $otherRelation['target']), $otherRelation['contents']);
                }
            }

            // Add presentation to ZIP file
            $objZip->addFromString('ppt/presentation.xml', $wPartPresentation->writePresentation($this->presentation));

            // Add slides (drawings, ...) and slide relationships (drawings, ...)
            for ($i = 0; $i < $this->presentation->getSlideCount(); ++$i) {
                // Add slide
                $objZip->addFromString('ppt/slides/_rels/slide' . ($i + 1) . '.xml.rels', $wPartRels->writeSlideRelationships($this->presentation->getSlide($i)));
                $objZip->addFromString('ppt/slides/slide' . ($i + 1) . '.xml', $wPartSlide->writeSlide($this->presentation->getSlide($i)));
            }

            // Add media
            for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
                $shape = $this->getDrawingHashTable()->getByIndex($i);
                if ($shape instanceof DrawingShape) {
                    $imagePath     = $shape->getPath();

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

                    $objZip->addFromString('ppt/media/' . str_replace(' ', '_', $shape->getIndexedFilename()), $imageContents);
                } elseif ($shape instanceof MemoryDrawingShape) {
                    ob_start();
                    call_user_func($shape->getRenderingFunction(), $shape->getImageResource());
                    $imageContents = ob_get_contents();
                    ob_end_clean();

                    $objZip->addFromString('ppt/media/' . str_replace(' ', '_', $shape->getIndexedFilename()), $imageContents);
                } elseif ($shape instanceof ChartShape) {
                    $objZip->addFromString('ppt/charts/' . $shape->getIndexedFilename(), $wPartChart->writeChart($shape));

                    // Chart relations?
                    if ($shape->hasIncludedSpreadsheet()) {
                        $objZip->addFromString('ppt/charts/_rels/' . $shape->getIndexedFilename() . '.rels', $wPartRels->writeChartRelationships($shape));
                        $objZip->addFromString('ppt/embeddings/' . $shape->getIndexedFilename() . '.xlsx', $wPartChart->writeSpreadsheet($this->presentation, $shape, $pFilename . '.xlsx'));
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
     * Get hash table
     *
     * @return \PhpOffice\PhpPowerpoint\HashTable
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
     * @return \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007
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
     * @return \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\AbstractLayoutPack
     */
    public function getLayoutPack()
    {
        return $this->layoutPack;
    }

    /**
     * Set layout pack to use
     *
     * @param \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\AbstractLayoutPack $pValue
     * @return \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007
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
