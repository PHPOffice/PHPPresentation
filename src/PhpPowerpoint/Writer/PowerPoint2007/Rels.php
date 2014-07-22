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

namespace PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\Chart as ShapeChart;
use PhpOffice\PhpPowerpoint\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;
use PhpOffice\PhpPowerpoint\Shape\RichText\Run;
use PhpOffice\PhpPowerpoint\Shape\RichText;
use PhpOffice\PhpPowerpoint\Shape\RichText\TextElement;
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;
use PhpOffice\PhpPowerpoint\Slide as SlideElement;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;

/**
 * \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Rels
 */
class Rels extends AbstractPart
{
    /**
     * Write relationships to XML format
     *
     * @return string        XML Output
     * @throws \Exception
     */
    public function writeRelationships()
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Relationship docProps/app.xml
        $this->writeRelationship($objWriter, 3, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'docProps/app.xml');

        // Relationship docProps/core.xml
        $this->writeRelationship($objWriter, 2, 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml');

        // Relationship ppt/presentation.xml
        $this->writeRelationship($objWriter, 1, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'ppt/presentation.xml');

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write presentation relationships to XML format
     *
     * @param  PHPPowerPoint $pPHPPowerPoint
     * @return string        XML Output
     * @throws \Exception
     */
    public function writePresentationRelationships(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Relation id
        $relationId = 1;

        $parentWriter = $this->getParentWriter();
        if ($parentWriter instanceof PowerPoint2007) {
            // Add slide masters
            $masterSlides = $parentWriter->getLayoutPack()->getMasterSlides();
            foreach ($masterSlides as $masterSlide) {
                // Relationship slideMasters/slideMasterX.xml
                $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', 'slideMasters/slideMaster' . $masterSlide['masterid'] . '.xml');
            }
        }

        // Add slide theme (only one!)
        // Relationship theme/theme1.xml
        $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml');

        // Relationships with slides
        $slideCount = $pPHPPowerPoint->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            $this->writeRelationship($objWriter, ($i + $relationId), 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slides/slide' . ($i + 1) . '.xml');
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write slide master relationships to XML format
     *
     * @param  int       $masterId Master slide id
     * @return string    XML Output
     * @throws \Exception
     */
    public function writeSlideMasterRelationships($masterId = 1)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Keep content id
        $contentId = 0;

        $parentWriter = $this->getParentWriter();
        if ($parentWriter instanceof PowerPoint2007) {
            // Lookup layouts
            $layouts    = array();
            $layoutPack = $parentWriter->getLayoutPack();
            foreach ($layoutPack->getLayouts() as $key => $layout) {
                if ($layout['masterid'] == $masterId) {
                    $layouts[$key] = $layout;
                }
            }
        
            // Write slideLayout relationships
            foreach ($layouts as $key => $layout) {
                $this->writeRelationship($objWriter, ++$contentId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout' . $key . '.xml');
            }
    
            // Relationship theme/theme1.xml
            $this->writeRelationship($objWriter, ++$contentId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', '../theme/theme' . $masterId . '.xml');
    
            // Other relationships
            $otherRelations = $layoutPack->getMasterSlideRelations();
            foreach ($otherRelations as $otherRelation) {
                if ($otherRelation['masterid'] == $masterId) {
                    $this->writeRelationship($objWriter, ++$contentId, $otherRelation['type'], $otherRelation['target']);
                }
            }
        }
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write slide layout relationships to XML format
     *
     * @param  int       $slideLayoutIndex
     * @param  int       $masterId
     * @return string    XML Output
     * @throws \Exception
     */
    public function writeSlideLayoutRelationships($slideLayoutIndex, $masterId = 1)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        $parentWriter = $this->getParentWriter();
        if ($parentWriter instanceof PowerPoint2007) {
            // Layout pack
            $layoutPack = $parentWriter->getLayoutPack();
    
            // XML header
            $objWriter->startDocument('1.0', 'UTF-8', 'yes');
    
            // Relationships
            $objWriter->startElement('Relationships');
            $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
    
            // Write slideMaster relationship
            $this->writeRelationship($objWriter, 1, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', '../slideMasters/slideMaster' . $masterId . '.xml');
    
            // Other relationships
            $otherRelations = $layoutPack->getLayoutRelations();
            foreach ($otherRelations as $otherRelation) {
                if ($otherRelation['layoutId'] == $slideLayoutIndex) {
                    $this->writeRelationship($objWriter, $otherRelation['id'], $otherRelation['type'], $otherRelation['target']);
                }
            }
    
            $objWriter->endElement();
        }

        // Return
        return $objWriter->getData();
    }

    /**
     * Write theme relationships to XML format
     *
     * @param  int       $masterId Master slide id
     * @return string    XML Output
     * @throws \Exception
     */
    public function writeThemeRelationships($masterId = 1)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        $parentWriter = $this->getParentWriter();
        if ($parentWriter instanceof PowerPoint2007) {
            // Layout pack
            $layoutPack = $parentWriter->getLayoutPack();
    
            // XML header
            $objWriter->startDocument('1.0', 'UTF-8', 'yes');
    
            // Relationships
            $objWriter->startElement('Relationships');
            $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
    
            // Other relationships
            $otherRelations = $layoutPack->getThemeRelations();
            foreach ($otherRelations as $otherRelation) {
                if ($otherRelation['masterid'] == $masterId) {
                    $this->writeRelationship($objWriter, $otherRelation['id'], $otherRelation['type'], $otherRelation['target']);
                }
            }
    
            $objWriter->endElement();
        }
        // Return
        return $objWriter->getData();
    }

    /**
     * Write slide relationships to XML format
     *
     * @param  \PhpOffice\PhpPowerpoint\Slide $pSlide
     * @return string              XML Output
     * @throws \Exception
     */
    public function writeSlideRelationships(SlideElement $pSlide)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Starting relation id
        $relId = 1;

        // Write slideLayout relationship
        $parentWriter = $this->getParentWriter();
        if ($parentWriter instanceof PowerPoint2007) {
            $layoutPack  = $parentWriter->getLayoutPack();
            $layoutIndex = $layoutPack->findlayoutIndex($pSlide->getSlideLayout(), $pSlide->getSlideMasterId());
    
            $this->writeRelationship($objWriter, $relId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout' . ($layoutIndex + 1) . '.xml');
        }

        // Write drawing relationships?
        if ($pSlide->getShapeCollection()->count() > 0) {
            // Loop trough images and write relationships
            $iterator = $pSlide->getShapeCollection()->getIterator();
            while ($iterator->valid()) {
                if ($iterator->current() instanceof ShapeDrawing || $iterator->current() instanceof MemoryDrawing) {
                    // Write relationship for image drawing
                    $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . str_replace(' ', '_', $iterator->current()->getIndexedFilename()));

                    $iterator->current()->relationId = 'rId' . $relId;

                    ++$relId;
                } elseif ($iterator->current() instanceof ShapeChart) {
                    // Write relationship for chart drawing
                    $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart', '../charts/' . $iterator->current()->getIndexedFilename());

                    $iterator->current()->relationId = 'rId' . $relId;

                    ++$relId;
                }

                $iterator->next();
            }
        }

        // Write hyperlink relationships?
        if ($pSlide->getShapeCollection()->count() > 0) {
            // Loop trough hyperlinks and write relationships
            $iterator = $pSlide->getShapeCollection()->getIterator();
            while ($iterator->valid()) {
                // Hyperlink on shape
                if ($iterator->current()->hasHyperlink()) {
                    // Write relationship for hyperlink
                    $hyperlink               = $iterator->current()->getHyperlink();
                    $hyperlink->relationId = 'rId' . $relId;

                    if (!$hyperlink->isInternal()) {
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                    } else {
                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                    }

                    ++$relId;
                }

                // Hyperlink on rich text run
                if ($iterator->current() instanceof RichText) {
                    foreach ($iterator->current()->getParagraphs() as $paragraph) {
                        foreach ($paragraph->getRichTextElements() as $element) {
                            if ($element instanceof Run || $element instanceof TextElement) {
                                if ($element->hasHyperlink()) {
                                    // Write relationship for hyperlink
                                    $hyperlink               = $element->getHyperlink();
                                    $hyperlink->relationId = 'rId' . $relId;

                                    if (!$hyperlink->isInternal()) {
                                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                                    } else {
                                        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                                    }

                                    ++$relId;
                                }
                            }
                        }
                    }
                }

                $iterator->next();
            }
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write chart relationships to XML format
     *
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart $pChart
     * @return string                    XML Output
     * @throws \Exception
     */
    public function writeChartRelationships(ShapeChart $pChart)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Write spreadsheet relationship?
        if ($pChart->hasIncludedSpreadsheet()) {
            $this->writeRelationship($objWriter, 1, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package', '../embeddings/' . $pChart->getIndexedFilename() . '.xlsx');
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write relationship
     *
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter $objWriter   XML Writer
     * @param  int                            $pId         Relationship ID. rId will be prepended!
     * @param  string                         $pType       Relationship type
     * @param  string                         $pTarget     Relationship target
     * @param  string                         $pTargetMode Relationship target mode
     * @throws \Exception
     */
    private function writeRelationship(XMLWriter $objWriter, $pId = 1, $pType = '', $pTarget = '', $pTargetMode = '')
    {
        if ($pType != '' && $pTarget != '') {
            if (strpos($pId, 'rId') === false) {
                $pId = 'rId' . $pId;
            }

            // Write relationship
            $objWriter->startElement('Relationship');
            $objWriter->writeAttribute('Id', $pId);
            $objWriter->writeAttribute('Type', $pType);
            $objWriter->writeAttribute('Target', $pTarget);

            if ($pTargetMode != '') {
                $objWriter->writeAttribute('TargetMode', $pTargetMode);
            }

            $objWriter->endElement();
        } else {
            throw new \Exception("Invalid parameters passed.");
        }
    }
}
