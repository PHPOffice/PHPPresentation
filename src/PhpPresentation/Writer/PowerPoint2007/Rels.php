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

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart as ShapeChart;
use PhpOffice\PhpPresentation\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table as ShapeTable;
use PhpOffice\PhpPresentation\Slide as SlideElement;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007\Rels
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
     * @param  PhpPresentation $pPhpPresentation
     * @return string        XML Output
     * @throws \Exception
     */
    public function writePresentationRelationships(PhpPresentation $pPhpPresentation)
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
        $slideCount = $pPhpPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slides/slide' . ($i + 1) . '.xml');
        }

        $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps', 'presProps.xml');
        $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps', 'viewProps.xml');
        $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles', 'tableStyles.xml');

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
     * @param  \PhpOffice\PhpPresentation\Slide $pSlide
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
        $idxSlide = $pSlide->getParent()->getIndex($pSlide);

        // Write slideLayout relationship
        $parentWriter = $this->getParentWriter();
        if ($parentWriter instanceof PowerPoint2007) {
            $layoutPack  = $parentWriter->getLayoutPack();
            $layoutId = $layoutPack->findlayoutId($pSlide->getSlideLayout(), $pSlide->getSlideMasterId());
    
            $this->writeRelationship($objWriter, $relId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout' . $layoutId . '.xml');
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
                } elseif ($iterator->current() instanceof Group) {
                    $iterator2 = $iterator->current()->getShapeCollection()->getIterator();
                    while ($iterator2->valid()) {
                        if ($iterator2->current() instanceof ShapeDrawing || $iterator2->current() instanceof MemoryDrawing) {
                            // Write relationship for image drawing
                            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . str_replace(' ', '_', $iterator2->current()->getIndexedFilename()));

                            $iterator2->current()->relationId = 'rId' . $relId;

                            ++$relId;
                        } elseif ($iterator2->current() instanceof ShapeChart) {
                            // Write relationship for chart drawing
                            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart', '../charts/' . $iterator2->current()->getIndexedFilename());

                            $iterator2->current()->relationId = 'rId' . $relId;

                            ++$relId;
                        }
                        $iterator2->next();
                    }
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
                
                // Hyperlink in table
                if ($iterator->current() instanceof ShapeTable) {
                    // Rows
                    $countRows = count($iterator->current()->getRows());
                    for ($row = 0; $row < $countRows; $row++) {
                        // Cells in rows
                        $countCells = count($iterator->current()->getRow($row)->getCells());
                        for ($cell = 0; $cell < $countCells; $cell++) {
                            $currentCell = $iterator->current()->getRow($row)->getCell($cell);
                            // Paragraphs in cell
                            foreach ($currentCell->getParagraphs() as $paragraph) {
                                // RichText in paragraph
                                foreach ($paragraph->getRichTextElements() as $element) {
                                    // Run or Text in RichText
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
                    }
                }

                if ($iterator->current() instanceof Group) {
                    $iterator2 = $pSlide->getShapeCollection()->getIterator();
                    while ($iterator2->valid()) {
                        // Hyperlink on shape
                        if ($iterator2->current()->hasHyperlink()) {
                            // Write relationship for hyperlink
                            $hyperlink             = $iterator2->current()->getHyperlink();
                            $hyperlink->relationId = 'rId' . $relId;

                            if (!$hyperlink->isInternal()) {
                                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', $hyperlink->getUrl(), 'External');
                            } else {
                                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slide' . $hyperlink->getSlideNumber() . '.xml');
                            }

                            ++$relId;
                        }

                        // Hyperlink on rich text run
                        if ($iterator2->current() instanceof RichText) {
                            foreach ($iterator2->current()->getParagraphs() as $paragraph) {
                                foreach ($paragraph->getRichTextElements() as $element) {
                                    if ($element instanceof Run || $element instanceof TextElement) {
                                        if ($element->hasHyperlink()) {
                                            // Write relationship for hyperlink
                                            $hyperlink              = $element->getHyperlink();
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
                        
                        // Hyperlink in table
                        if ($iterator2->current() instanceof ShapeTable) {
                            // Rows
                            $countRows = count($iterator2->current()->getRows());
                            for ($row = 0; $row < $countRows; $row++) {
                                // Cells in rows
                                $countCells = count($iterator2->current()->getRow($row)->getCells());
                                for ($cell = 0; $cell < $countCells; $cell++) {
                                    $currentCell = $iterator2->current()->getRow($row)->getCell($cell);
                                    // Paragraphs in cell
                                    foreach ($currentCell->getParagraphs() as $paragraph) {
                                        // RichText in paragraph
                                        foreach ($paragraph->getRichTextElements() as $element) {
                                            // Run or Text in RichText
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
                            }
                        }
                        
                        $iterator2->next();
                    }
                    
                }

                $iterator->next();
            }
        }
        
        if ($pSlide->getNote()->getShapeCollection()->count() > 0) {
            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide', '../notesSlides/notesSlide'.($idxSlide + 1).'.xml');
            ++$relId;
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write chart relationships to XML format
     *
     * @param  \PhpOffice\PhpPresentation\Shape\Chart $pChart
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
     * @param  \PhpOffice\Common\XMLWriter $objWriter   XML Writer
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
