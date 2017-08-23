<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Chart as ShapeChart;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table as ShapeTable;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Slide\Note;

class PptSlides extends AbstractSlide
{
    /**
     * Add slides (drawings, ...) and slide relationships (drawings, ...)
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        foreach ($this->oPresentation->getAllSlides() as $idx => $oSlide) {
            $this->oZip->addFromString('ppt/slides/_rels/slide' . ($idx + 1) . '.xml.rels', $this->writeSlideRelationships($oSlide));
            $this->oZip->addFromString('ppt/slides/slide' . ($idx + 1) . '.xml', $this->writeSlide($oSlide));

            // Add note slide
            if ($oSlide->getNote() instanceof Note) {
                if ($oSlide->getNote()->getShapeCollection()->count() > 0) {
                    $this->oZip->addFromString('ppt/notesSlides/notesSlide' . ($idx + 1) . '.xml', $this->writeNote($oSlide->getNote()));
                }
            }

            // Add background image slide
            $oBkgImage = $oSlide->getBackground();
            if ($oBkgImage instanceof Image) {
                $this->oZip->addFromString('ppt/media/'.$oBkgImage->getIndexedFilename($idx), file_get_contents($oBkgImage->getPath()));
            }
        }

        return $this->oZip;
    }

    /**
     * Write slide relationships to XML format
     *
     * @param  \PhpOffice\PhpPresentation\Slide $pSlide
     * @return string              XML Output
     * @throws \Exception
     */
    protected function writeSlideRelationships(Slide $pSlide)
    {
        //@todo Group all getShapeCollection()->getIterator

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Starting relation id
        $relId = 1;
        $idxSlide = $pSlide->getParent()->getIndex($pSlide);

        // Write slideLayout relationship
        $layoutId = 1;
        if ($pSlide->getSlideLayout()) {
            $layoutId = $pSlide->getSlideLayout()->layoutNr;
        }
        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout' . $layoutId . '.xml');
        ++$relId;

        // Write drawing relationships?
        if ($pSlide->getShapeCollection()->count() > 0) {
            // Loop trough images and write relationships
            $iterator = $pSlide->getShapeCollection()->getIterator();
            while ($iterator->valid()) {
                if ($iterator->current() instanceof Media) {
                    // Write relationship for image drawing
                    $iterator->current()->relationId = 'rId' . $relId;
                    $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video', '../media/' . $iterator->current()->getIndexedFilename());
                    ++$relId;
                    $this->writeRelationship($objWriter, $relId, 'http://schemas.microsoft.com/office/2007/relationships/media', '../media/' . $iterator->current()->getIndexedFilename());
                    ++$relId;
                } elseif ($iterator->current() instanceof ShapeDrawing\AbstractDrawingAdapter) {
                    // Write relationship for image drawing
                    $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $iterator->current()->getIndexedFilename());
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
                        if ($iterator2->current() instanceof Media) {
                            // Write relationship for image drawing
                            $iterator2->current()->relationId = 'rId' . $relId;
                            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video', '../media/' . $iterator->current()->getIndexedFilename());
                            ++$relId;
                            $this->writeRelationship($objWriter, $relId, 'http://schemas.microsoft.com/office/2007/relationships/media', '../media/' . $iterator->current()->getIndexedFilename());
                            ++$relId;
                        } elseif ($iterator2->current() instanceof ShapeDrawing\AbstractDrawingAdapter) {
                            // Write relationship for image drawing
                            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $iterator2->current()->getIndexedFilename());
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

        // Write background relationships?
        $oBackground = $pSlide->getBackground();
        if ($oBackground instanceof Image) {
            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $oBackground->getIndexedFilename($idxSlide));
            $oBackground->relationId = 'rId' . $relId;
            ++$relId;
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

        // Write comment relationships
        if ($pSlide->getShapeCollection()->count() > 0) {
            $hasSlideComment = false;

            // Loop trough images and write relationships
            $iterator = $pSlide->getShapeCollection()->getIterator();
            while ($iterator->valid()) {
                if ($iterator->current() instanceof Comment) {
                    $hasSlideComment = true;
                    break;
                } elseif ($iterator->current() instanceof Group) {
                    $iterator2 = $iterator->current()->getShapeCollection()->getIterator();
                    while ($iterator2->valid()) {
                        if ($iterator2->current() instanceof Comment) {
                            $hasSlideComment = true;
                            break 2;
                        }
                        $iterator2->next();
                    }
                }

                $iterator->next();
            }

            if ($hasSlideComment) {
                $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', '../comments/comment'.($idxSlide + 1).'.xml');
                ++$relId;
            }
        }

        if ($pSlide->getNote()->getShapeCollection()->count() > 0) {
            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide', '../notesSlides/notesSlide'.($idxSlide + 1).'.xml');
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write slide to XML format
     *
     * @param  \PhpOffice\PhpPresentation\Slide $pSlide
     * @return string              XML Output
     * @throws \Exception
     */
    public function writeSlide(Slide $pSlide)
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:sld
        $objWriter->startElement('p:sld');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $objWriter->writeAttributeIf(!$pSlide->isVisible(), 'show', 0);

        // p:sld/p:cSld
        $objWriter->startElement('p:cSld');

        // Background
        if ($pSlide->getBackground() instanceof Slide\AbstractBackground) {
            $oBackground = $pSlide->getBackground();
            // p:bg
            $objWriter->startElement('p:bg');

            // p:bgPr
            $objWriter->startElement('p:bgPr');

            if ($oBackground instanceof Slide\Background\Color) {
                // a:solidFill
                $objWriter->startElement('a:solidFill');

                $this->writeColor($objWriter, $oBackground->getColor());

                // > a:solidFill
                $objWriter->endElement();
            }

            if ($oBackground instanceof Slide\Background\Image) {
                // a:blipFill
                $objWriter->startElement('a:blipFill');

                // a:blip
                $objWriter->startElement('a:blip');
                $objWriter->writeAttribute('r:embed', $oBackground->relationId);

                // > a:blipFill
                $objWriter->endElement();

                // a:stretch
                $objWriter->startElement('a:stretch');

                // a:fillRect
                $objWriter->writeElement('a:fillRect');

                // > a:stretch
                $objWriter->endElement();

                // > a:blipFill
                $objWriter->endElement();
            }

            // > p:bgPr
            $objWriter->endElement();

            // > p:bg
            $objWriter->endElement();
        }

        // p:spTree
        $objWriter->startElement('p:spTree');

        // p:nvGrpSpPr
        $objWriter->startElement('p:nvGrpSpPr');

        // p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '1');
        $objWriter->writeAttribute('name', '');
        $objWriter->endElement();

        // p:cNvGrpSpPr
        $objWriter->writeElement('p:cNvGrpSpPr', null);

        // p:nvPr
        $objWriter->writeElement('p:nvPr', null);

        $objWriter->endElement();

        // p:grpSpPr
        $objWriter->startElement('p:grpSpPr');

        // a:xfrm
        $objWriter->startElement('a:xfrm');

        // a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pSlide->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pSlide->getOffsetY()));
        $objWriter->endElement(); // a:off

        // a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pSlide->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pSlide->getExtentY()));
        $objWriter->endElement(); // a:ext

        // a:chOff
        $objWriter->startElement('a:chOff');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pSlide->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pSlide->getOffsetY()));
        $objWriter->endElement(); // a:chOff

        // a:chExt
        $objWriter->startElement('a:chExt');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pSlide->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pSlide->getExtentY()));
        $objWriter->endElement(); // a:chExt

        $objWriter->endElement();

        $objWriter->endElement();

        // Loop shapes
        $this->writeShapeCollection($objWriter, $pSlide->getShapeCollection());

        // TODO
        $objWriter->endElement();

        $objWriter->endElement();

        // p:clrMapOvr
        $objWriter->startElement('p:clrMapOvr');
        // p:clrMapOvr\a:masterClrMapping
        $objWriter->writeElement('a:masterClrMapping', null);
        // ##p:clrMapOvr
        $objWriter->endElement();

        $this->writeSlideTransition($objWriter, $pSlide->getTransition());

        $this->writeSlideAnimations($objWriter, $pSlide);

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * @param XMLWriter $objWriter
     * @param Slide $oSlide
     */
    protected function writeSlideAnimations(XMLWriter $objWriter, Slide $oSlide)
    {
        $arrayAnimations = $oSlide->getAnimations();
        if (empty($arrayAnimations)) {
            return;
        }

        // Variables
        $shapeId = 1;
        $idCount = 1;
        $hashToIdMap = array();
        $arrayAnimationIds = array();

        foreach ($oSlide->getShapeCollection() as $shape) {
            $hashToIdMap[$shape->getHashCode()] = ++$shapeId;
        }
        foreach ($arrayAnimations as $oAnimation) {
            foreach ($oAnimation->getShapeCollection() as $oShape) {
                $arrayAnimationIds[] = $hashToIdMap[$oShape->getHashCode()];
            }
        }

        // p:timing
        $objWriter->startElement('p:timing');
        // p:timing/p:tnLst
        $objWriter->startElement('p:tnLst');
        // p:timing/p:tnLst/p:par
        $objWriter->startElement('p:par');
        // p:timing/p:tnLst/p:par/p:cTn
        $objWriter->startElement('p:cTn');
        $objWriter->writeAttribute('id', $idCount++);
        $objWriter->writeAttribute('dur', 'indefinite');
        $objWriter->writeAttribute('restart', 'never');
        $objWriter->writeAttribute('nodeType', 'tmRoot');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst
        $objWriter->startElement('p:childTnLst');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq
        $objWriter->startElement('p:seq');
        $objWriter->writeAttribute('concurrent', '1');
        $objWriter->writeAttribute('nextAc', 'seek');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn
        $objWriter->startElement('p:cTn');
        $objWriter->writeAttribute('id', $idCount++);
        $objWriter->writeAttribute('dur', 'indefinite');
        $objWriter->writeAttribute('nodeType', 'mainSeq');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst
        $objWriter->startElement('p:childTnLst');

        // Each animation has multiple shapes
        foreach ($arrayAnimations as $oAnimation) {
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par
            $objWriter->startElement('p:par');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst/p:cond
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', 'indefinite');
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn\##p:stCondLst
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst
            $objWriter->startElement('p:childTnLst');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par
            $objWriter->startElement('p:par');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn
            $objWriter->startElement('p:cTn');
            $objWriter->writeAttribute('id', $idCount++);
            $objWriter->writeAttribute('fill', 'hold');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst
            $objWriter->startElement('p:stCondLst');
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:stCondLst/p:cond
            $objWriter->startElement('p:cond');
            $objWriter->writeAttribute('delay', '0');
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn\##p:stCondLst
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst
            $objWriter->startElement('p:childTnLst');

            $firstAnimation = true;
            foreach ($oAnimation->getShapeCollection() as $oShape) {
                $nodeType = $firstAnimation ? 'clickEffect' : 'withEffect';
                $shapeId = $hashToIdMap[$oShape->getHashCode()];

                // p:par
                $objWriter->startElement('p:par');
                // p:par/p:cTn
                $objWriter->startElement('p:cTn');
                $objWriter->writeAttribute('id', $idCount++);
                $objWriter->writeAttribute('presetID', '1');
                $objWriter->writeAttribute('presetClass', 'entr');
                $objWriter->writeAttribute('fill', 'hold');
                $objWriter->writeAttribute('presetSubtype', '0');
                $objWriter->writeAttribute('grpId', '0');
                $objWriter->writeAttribute('nodeType', $nodeType);
                // p:par/p:cTn/p:stCondLst
                $objWriter->startElement('p:stCondLst');
                // p:par/p:cTn/p:stCondLst/p:cond
                $objWriter->startElement('p:cond');
                $objWriter->writeAttribute('delay', '0');
                $objWriter->endElement();
                // p:par/p:cTn\##p:stCondLst
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst
                $objWriter->startElement('p:childTnLst');
                // p:par/p:cTn/p:childTnLst/p:set
                $objWriter->startElement('p:set');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr
                $objWriter->startElement('p:cBhvr');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:cTn
                $objWriter->startElement('p:cTn');
                $objWriter->writeAttribute('id', $idCount++);
                $objWriter->writeAttribute('dur', '1');
                $objWriter->writeAttribute('fill', 'hold');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:cTn/p:stCondLst
                $objWriter->startElement('p:stCondLst');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:cTn/p:stCondLst/p:cond
                $objWriter->startElement('p:cond');
                $objWriter->writeAttribute('delay', '0');
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:cTn\##p:stCondLst
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr\##p:cTn
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:tgtEl
                $objWriter->startElement('p:tgtEl');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:tgtEl/p:spTgt
                $objWriter->startElement('p:spTgt');
                $objWriter->writeAttribute('spid', $shapeId);
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr\##p:tgtEl
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:attrNameLst
                $objWriter->startElement('p:attrNameLst');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr/p:attrNameLst/p:attrName
                $objWriter->writeElement('p:attrName', 'style.visibility');
                // p:par/p:cTn/p:childTnLst/p:set/p:cBhvr\##p:attrNameLst
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set\##p:cBhvr
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set/p:to
                $objWriter->startElement('p:to');
                // p:par/p:cTn/p:childTnLst/p:set/p:to/p:strVal
                $objWriter->startElement('p:strVal');
                $objWriter->writeAttribute('val', 'visible');
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst/p:set\##p:to
                $objWriter->endElement();
                // p:par/p:cTn/p:childTnLst\##p:set
                $objWriter->endElement();
                // p:par/p:cTn\##p:childTnLst
                $objWriter->endElement();
                // p:par\##p:cTn
                $objWriter->endElement();
                // ##p:par
                $objWriter->endElement();

                $firstAnimation = false;
            }

            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn\##p:childTnLst
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst/p:par\##p:cTn
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn/p:childTnLst\##p:par
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par/p:cTn\##p:childTnLst
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst/p:par\##p:cTn
            $objWriter->endElement();
            // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn/p:childTnLst\##p:par
            $objWriter->endElement();
        }

        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:cTn\##p:childTnLst
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq\##p:cTn
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst
        $objWriter->startElement('p:prevCondLst');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst/p:cond
        $objWriter->startElement('p:cond');
        $objWriter->writeAttribute('evt', 'onPrev');
        $objWriter->writeAttribute('delay', '0');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst/p:cond/p:tgtEl
        $objWriter->startElement('p:tgtEl');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst/p:cond/p:tgtEl/p:sldTgt
        $objWriter->writeElement('p:sldTgt', null);
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst/p:cond\##p:tgtEl
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:prevCondLst\##p:cond
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq\##p:prevCondLst
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst
        $objWriter->startElement('p:nextCondLst');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst/p:cond
        $objWriter->startElement('p:cond');
        $objWriter->writeAttribute('evt', 'onNext');
        $objWriter->writeAttribute('delay', '0');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst/p:cond/p:tgtEl
        $objWriter->startElement('p:tgtEl');
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst/p:cond/p:tgtEl/p:sldTgt
        $objWriter->writeElement('p:sldTgt', null);
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst/p:cond\##p:tgtEl
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq/p:nextCondLst\##p:cond
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst/p:seq\##p:nextCondLst
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn/p:childTnLst\##p:seq
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par/p:cTn\##p:childTnLst
        $objWriter->endElement();
        // p:timing/p:tnLst/p:par\##p:cTn
        $objWriter->endElement();
        // p:timing/p:tnLst\##p:par
        $objWriter->endElement();
        // p:timing\##p:tnLst
        $objWriter->endElement();

        // p:timing/p:bldLst
        $objWriter->startElement('p:bldLst');

        // Add in ids of all shapes in this slides animations
        foreach ($arrayAnimationIds as $id) {
            // p:timing/p:bldLst/p:bldP
            $objWriter->startElement('p:bldP');
            $objWriter->writeAttribute('spid', $id);
            $objWriter->writeAttribute('grpId', 0);
            $objWriter->endELement();
        }

        // p:timing\##p:bldLst
        $objWriter->endElement();

        // ##p:timing
        $objWriter->endElement();
    }

    /**
     * Write pic
     *
     * @param  \PhpOffice\Common\XMLWriter  $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter $shape
     * @param  int $shapeId
     * @throws \Exception
     */
    protected function writeShapeDrawing(XMLWriter $objWriter, ShapeDrawing\AbstractDrawingAdapter $shape, $shapeId)
    {
        // p:pic
        $objWriter->startElement('p:pic');

        // p:nvPicPr
        $objWriter->startElement('p:nvPicPr');

        // p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', $shapeId);
        $objWriter->writeAttribute('name', $shape->getName());
        $objWriter->writeAttribute('descr', $shape->getDescription());

        // a:hlinkClick
        if ($shape->hasHyperlink()) {
            $this->writeHyperlink($objWriter, $shape);
        }

        $objWriter->endElement();

        // p:cNvPicPr
        $objWriter->startElement('p:cNvPicPr');

        // a:picLocks
        $objWriter->startElement('a:picLocks');
        $objWriter->writeAttribute('noChangeAspect', '1');
        $objWriter->endElement();

        $objWriter->endElement();

        // p:nvPr
        $objWriter->startElement('p:nvPr');
        // PlaceHolder
        if ($shape->isPlaceholder()) {
            $objWriter->startElement('p:ph');
            $objWriter->writeAttribute('type', $shape->getPlaceholder()->getType());
            $objWriter->endElement();
        }
        /**
         * @link : https://github.com/stefslon/exportToPPTX/blob/master/exportToPPTX.m#L2128
         */
        if ($shape instanceof Media) {
            // p:nvPr > a:videoFile
            $objWriter->startElement('a:videoFile');
            $objWriter->writeAttribute('r:link', $shape->relationId);
            $objWriter->endElement();
            // p:nvPr > p:extLst
            $objWriter->startElement('p:extLst');
            // p:nvPr > p:extLst > p:ext
            $objWriter->startElement('p:ext');
            $objWriter->writeAttribute('uri', '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}');
            // p:nvPr > p:extLst > p:ext > p14:media
            $objWriter->startElement('p14:media');
            $objWriter->writeAttribute('r:embed', ($shape->relationId + 1));
            $objWriter->writeAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
            // p:nvPr > p:extLst > p:ext > ##p14:media
            $objWriter->endElement();
            // p:nvPr > p:extLst > ##p:ext
            $objWriter->endElement();
            // p:nvPr > ##p:extLst
            $objWriter->endElement();
        }
        // ##p:nvPr
        $objWriter->endElement();
        $objWriter->endElement();

        // p:blipFill
        $objWriter->startElement('p:blipFill');

        // a:blip
        $objWriter->startElement('a:blip');
        $objWriter->writeAttribute('r:embed', $shape->relationId);
        $objWriter->endElement();

        // a:stretch
        $objWriter->startElement('a:stretch');
        $objWriter->writeElement('a:fillRect', null);
        $objWriter->endElement();

        $objWriter->endElement();

        // p:spPr
        $objWriter->startElement('p:spPr');
        // a:xfrm
        $objWriter->startElement('a:xfrm');
        $objWriter->writeAttributeIf($shape->getRotation() != 0, 'rot', CommonDrawing::degreesToAngle($shape->getRotation()));

        // a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($shape->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($shape->getOffsetY()));
        $objWriter->endElement();

        // a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($shape->getWidth()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($shape->getHeight()));
        $objWriter->endElement();

        $objWriter->endElement();

        // a:prstGeom
        $objWriter->startElement('a:prstGeom');
        $objWriter->writeAttribute('prst', 'rect');

        // a:avLst
        $objWriter->writeElement('a:avLst', null);

        $objWriter->endElement();

        $this->writeBorder($objWriter, $shape->getBorder(), '');

        $this->writeShadow($objWriter, $shape->getShadow());

        $objWriter->endElement();

        $objWriter->endElement();
    }
}
