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

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\AbstractGraphic;
use PhpOffice\PhpPresentation\Shape\Chart as ShapeChart;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd as ShapeDrawingGd;
use PhpOffice\PhpPresentation\Shape\Drawing\File as ShapeDrawingFile;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table as ShapeTable;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Note;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Slide\AbstractSlide as AbstractSlideAlias;

abstract class AbstractSlide extends AbstractDecoratorWriter
{
    /**
     * @param AbstractSlideAlias $pSlideMaster
     * @param $objWriter
     * @param $relId
     * @throws \Exception
     */
    protected function writeDrawingRelations(AbstractSlideAlias $pSlideMaster, $objWriter, $relId)
    {
        if ($pSlideMaster->getShapeCollection()->count() > 0) {
            // Loop trough images and write relationships
            $iterator = $pSlideMaster->getShapeCollection()->getIterator();
            while ($iterator->valid()) {
                if ($iterator->current() instanceof ShapeDrawingFile || $iterator->current() instanceof ShapeDrawingGd) {
                    // Write relationship for image drawing
                    $this->writeRelationship(
                        $objWriter,
                        $relId,
                        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                        '../media/' . str_replace(' ', '_', $iterator->current()->getIndexedFilename())
                    );
                    $iterator->current()->relationId = 'rId' . $relId;
                    ++$relId;
                } elseif ($iterator->current() instanceof ShapeChart) {
                    // Write relationship for chart drawing
                    $this->writeRelationship(
                        $objWriter,
                        $relId,
                        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                        '../charts/' . $iterator->current()->getIndexedFilename()
                    );
                    $iterator->current()->relationId = 'rId' . $relId;
                    ++$relId;
                } elseif ($iterator->current() instanceof Group) {
                    $iterator2 = $iterator->current()->getShapeCollection()->getIterator();
                    while ($iterator2->valid()) {
                        if ($iterator2->current() instanceof ShapeDrawingFile ||
                            $iterator2->current() instanceof ShapeDrawingGd
                        ) {
                            // Write relationship for image drawing
                            $this->writeRelationship(
                                $objWriter,
                                $relId,
                                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                '../media/' . str_replace(' ', '_', $iterator2->current()->getIndexedFilename())
                            );
                            $iterator2->current()->relationId = 'rId' . $relId;
                            ++$relId;
                        } elseif ($iterator2->current() instanceof ShapeChart) {
                            // Write relationship for chart drawing
                            $this->writeRelationship(
                                $objWriter,
                                $relId,
                                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                                '../charts/' . $iterator2->current()->getIndexedFilename()
                            );
                            $iterator2->current()->relationId = 'rId' . $relId;
                            ++$relId;
                        }
                        $iterator2->next();
                    }
                }
                $iterator->next();
            }
        }

        return $relId;
    }

    /**
     * @param XMLWriter $objWriter
     * @param \ArrayObject|\PhpOffice\PhpPresentation\AbstractShape[] $shapes
     * @param int $shapeId
     * @throws \Exception
     */
    protected function writeShapeCollection(XMLWriter $objWriter, $shapes = array(), &$shapeId = 0)
    {
        if (count($shapes) == 0) {
            return;
        }
        foreach ($shapes as $shape) {
            // Increment $shapeId
            ++$shapeId;
            // Check type
            if ($shape instanceof RichText) {
                $this->writeShapeText($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof ShapeTable) {
                $this->writeShapeTable($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Line) {
                $this->writeShapeLine($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof ShapeChart) {
                $this->writeShapeChart($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof AbstractGraphic) {
                $this->writeShapePic($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Group) {
                $this->writeShapeGroup($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Comment) {
            } else {
                throw new \Exception("Unknown Shape type: {get_class($shape)}");
            }
        }
    }

    /**
     * Write txt
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\RichText $shape
     * @param  int $shapeId
     * @throws \Exception
     */
    protected function writeShapeText(XMLWriter $objWriter, RichText $shape, $shapeId)
    {
        // p:sp
        $objWriter->startElement('p:sp');
        // p:sp\p:nvSpPr
        $objWriter->startElement('p:nvSpPr');
        // p:sp\p:nvSpPr\p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', $shapeId);
        if ($shape->isPlaceholder()) {
            $objWriter->writeAttribute('name', 'Placeholder for ' . $shape->getPlaceholder()->getType());
        } else {
            $objWriter->writeAttribute('name', '');
        }
        // Hyperlink
        if ($shape->hasHyperlink()) {
            $this->writeHyperlink($objWriter, $shape);
        }
        // > p:sp\p:nvSpPr
        $objWriter->endElement();
        // p:sp\p:cNvSpPr
        $objWriter->startElement('p:cNvSpPr');
        $objWriter->writeAttribute('txBox', '1');
        $objWriter->endElement();
        // p:sp\p:cNvSpPr\p:nvPr
        if ($shape->isPlaceholder()) {
            $objWriter->startElement('p:nvPr');
            $objWriter->startElement('p:ph');
            $objWriter->writeAttribute('type', $shape->getPlaceholder()->getType());
            if (!is_null($shape->getPlaceholder()->getIdx())) {
                $objWriter->writeAttribute('idx', $shape->getPlaceholder()->getIdx());
            }
            $objWriter->endElement();
            $objWriter->endElement();
        } else {
            $objWriter->writeElement('p:nvPr', null);
        }
        // > p:sp\p:cNvSpPr
        $objWriter->endElement();
        // p:sp\p:spPr
        $objWriter->startElement('p:spPr');

        if (!$shape->isPlaceholder()) {
            // p:sp\p:spPr\a:xfrm
            $objWriter->startElement('a:xfrm');
            $objWriter->writeAttributeIf($shape->getRotation() != 0, 'rot', CommonDrawing::degreesToAngle($shape->getRotation()));
            // p:sp\p:spPr\a:xfrm\a:off
            $objWriter->startElement('a:off');
            $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($shape->getOffsetX()));
            $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($shape->getOffsetY()));
            $objWriter->endElement();
            // p:sp\p:spPr\a:xfrm\a:ext
            $objWriter->startElement('a:ext');
            $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($shape->getWidth()));
            $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($shape->getHeight()));
            $objWriter->endElement();
            // > p:sp\p:spPr\a:xfrm
            $objWriter->endElement();
            // p:sp\p:spPr\a:prstGeom
            $objWriter->startElement('a:prstGeom');
            $objWriter->writeAttribute('prst', 'rect');

            // p:sp\p:spPr\a:prstGeom\a:avLst
            $objWriter->writeElement('a:avLst');

            $objWriter->endElement();
        }
        $this->writeFill($objWriter, $shape->getFill());
        $this->writeBorder($objWriter, $shape->getBorder(), '');
        $this->writeShadow($objWriter, $shape->getShadow());

        // > p:sp\p:spPr
        $objWriter->endElement();
        // p:txBody
        $objWriter->startElement('p:txBody');
        // a:bodyPr
        //@link :http://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.bodyproperties%28v=office.14%29.aspx
        $objWriter->startElement('a:bodyPr');
        if (!$shape->isPlaceholder()) {
            $verticalAlign = $shape->getActiveParagraph()->getAlignment()->getVertical();
            if ($verticalAlign != Alignment::VERTICAL_BASE && $verticalAlign != Alignment::VERTICAL_AUTO) {
                $objWriter->writeAttribute('anchor', $verticalAlign);
            }
            if ($shape->getWrap() != RichText::WRAP_SQUARE) {
                $objWriter->writeAttribute('wrap', $shape->getWrap());
            }
            $objWriter->writeAttribute('rtlCol', '0');
            if ($shape->getHorizontalOverflow() != RichText::OVERFLOW_OVERFLOW) {
                $objWriter->writeAttribute('horzOverflow', $shape->getHorizontalOverflow());
            }
            if ($shape->getVerticalOverflow() != RichText::OVERFLOW_OVERFLOW) {
                $objWriter->writeAttribute('vertOverflow', $shape->getVerticalOverflow());
            }
            if ($shape->isUpright()) {
                $objWriter->writeAttribute('upright', '1');
            }
            if ($shape->isVertical()) {
                $objWriter->writeAttribute('vert', 'vert');
            }
            $objWriter->writeAttribute('bIns', CommonDrawing::pixelsToEmu($shape->getInsetBottom()));
            $objWriter->writeAttribute('lIns', CommonDrawing::pixelsToEmu($shape->getInsetLeft()));
            $objWriter->writeAttribute('rIns', CommonDrawing::pixelsToEmu($shape->getInsetRight()));
            $objWriter->writeAttribute('tIns', CommonDrawing::pixelsToEmu($shape->getInsetTop()));
            if ($shape->getColumns() <> 1) {
                $objWriter->writeAttribute('numCol', $shape->getColumns());
            }
            // a:spAutoFit
            $objWriter->startElement('a:' . $shape->getAutoFit());
            if ($shape->getAutoFit() == RichText::AUTOFIT_NORMAL) {
                if (!is_null($shape->getFontScale())) {
                    $objWriter->writeAttribute('fontScale', (int)($shape->getFontScale() * 1000));
                }
                if (!is_null($shape->getLineSpaceReduction())) {
                    $objWriter->writeAttribute('lnSpcReduction', (int)($shape->getLineSpaceReduction() * 1000));
                }
            }
            $objWriter->endElement();
        }
        $objWriter->endElement();
        // a:lstStyle
        $objWriter->writeElement('a:lstStyle', null);
        if ($shape->isPlaceholder() &&
            ($shape->getPlaceholder()->getType() == Placeholder::PH_TYPE_SLIDENUM ||
                $shape->getPlaceholder()->getType() == Placeholder::PH_TYPE_DATETIME)
        ) {
            $objWriter->startElement('a:p');
            $objWriter->startElement('a:fld');
            $objWriter->writeAttribute('id', $this->getGUID());
            $objWriter->writeAttribute('type', (
            $shape->getPlaceholder()->getType() == Placeholder::PH_TYPE_SLIDENUM ? 'slidenum' : 'datetime'));
            $objWriter->writeElement('a:t', (
            $shape->getPlaceholder()->getType() == Placeholder::PH_TYPE_SLIDENUM ? '<nr.>' : '03-04-05'));
            $objWriter->endElement();
            $objWriter->endElement();
        } else {
            // Write paragraphs
            $this->writeParagraphs($objWriter, $shape->getParagraphs());
        }
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Write table
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Table $shape
     * @param  int $shapeId
     * @throws \Exception
     */
    protected function writeShapeTable(XMLWriter $objWriter, ShapeTable $shape, $shapeId)
    {
        // p:graphicFrame
        $objWriter->startElement('p:graphicFrame');
        // p:graphicFrame/p:nvGraphicFramePr
        $objWriter->startElement('p:nvGraphicFramePr');
        // p:graphicFrame/p:nvGraphicFramePr/p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', $shapeId);
        $objWriter->writeAttribute('name', $shape->getName());
        $objWriter->writeAttribute('descr', $shape->getDescription());
        $objWriter->endElement();
        // p:graphicFrame/p:nvGraphicFramePr/p:cNvGraphicFramePr
        $objWriter->startElement('p:cNvGraphicFramePr');
        // p:graphicFrame/p:nvGraphicFramePr/p:cNvGraphicFramePr/a:graphicFrameLocks
        $objWriter->startElement('a:graphicFrameLocks');
        $objWriter->writeAttribute('noGrp', '1');
        $objWriter->endElement();
        // p:graphicFrame/p:nvGraphicFramePr/p:cNvGraphicFramePr/
        $objWriter->endElement();
        // p:graphicFrame/p:nvGraphicFramePr/p:nvPr
        $objWriter->startElement('p:nvPr');
        if ($shape->isPlaceholder()) {
            $objWriter->startElement('p:ph');
            $objWriter->writeAttribute('type', $shape->getPlaceholder()->getType());
            $objWriter->endElement();
        }
        $objWriter->endElement();
        // p:graphicFrame/p:nvGraphicFramePr/
        $objWriter->endElement();
        // p:graphicFrame/p:xfrm
        $objWriter->startElement('p:xfrm');
        // p:graphicFrame/p:xfrm/a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($shape->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($shape->getOffsetY()));
        $objWriter->endElement();
        // p:graphicFrame/p:xfrm/a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($shape->getWidth()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($shape->getHeight()));
        $objWriter->endElement();
        // p:graphicFrame/p:xfrm/
        $objWriter->endElement();
        // p:graphicFrame/a:graphic
        $objWriter->startElement('a:graphic');
        // p:graphicFrame/a:graphic/a:graphicData
        $objWriter->startElement('a:graphicData');
        $objWriter->writeAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/table');
        // p:graphicFrame/a:graphic/a:graphicData/a:tbl
        $objWriter->startElement('a:tbl');
        // p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tblPr
        $objWriter->startElement('a:tblPr');
        $objWriter->writeAttribute('firstRow', '1');
        $objWriter->writeAttribute('bandRow', '1');
        $objWriter->endElement();
        // p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tblGrid
        $objWriter->startElement('a:tblGrid');
        // Write cell widths
        $countCells = count($shape->getRow(0)->getCells());
        for ($cell = 0; $cell < $countCells; $cell++) {
            //  p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tblGrid/a:gridCol
            $objWriter->startElement('a:gridCol');
            // Calculate column width
            $width = $shape->getRow(0)->getCell($cell)->getWidth();
            if ($width == 0) {
                $colCount = count($shape->getRow(0)->getCells());
                $totalWidth = $shape->getWidth();
                $width = $totalWidth / $colCount;
            }
            $objWriter->writeAttribute('w', CommonDrawing::pixelsToEmu($width));
            $objWriter->endElement();
        }
        // p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tblGrid/
        $objWriter->endElement();
        // Colspan / rowspan containers
        $colSpan = array();
        $rowSpan = array();
        // Default border style
        $defaultBorder = new Border();
        // Write rows
        $countRows = count($shape->getRows());
        for ($row = 0; $row < $countRows; $row++) {
            // p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr
            $objWriter->startElement('a:tr');
            $objWriter->writeAttribute('h', CommonDrawing::pixelsToEmu($shape->getRow($row)->getHeight()));
            // Write cells
            $countCells = count($shape->getRow($row)->getCells());
            for ($cell = 0; $cell < $countCells; $cell++) {
                // Current cell
                $currentCell = $shape->getRow($row)->getCell($cell);
                // Next cell right
                $nextCellRight = $shape->getRow($row)->getCell($cell + 1, true);
                // Next cell below
                $nextRowBelow = $shape->getRow($row + 1, true);
                $nextCellBelow = null;
                if ($nextRowBelow != null) {
                    $nextCellBelow = $nextRowBelow->getCell($cell, true);
                }
                // a:tc
                $objWriter->startElement('a:tc');
                // Colspan
                if ($currentCell->getColSpan() > 1) {
                    $objWriter->writeAttribute('gridSpan', $currentCell->getColSpan());
                    $colSpan[$row] = $currentCell->getColSpan() - 1;
                } elseif (isset($colSpan[$row]) && $colSpan[$row] > 0) {
                    $colSpan[$row]--;
                    $objWriter->writeAttribute('hMerge', '1');
                }
                // Rowspan
                if ($currentCell->getRowSpan() > 1) {
                    $objWriter->writeAttribute('rowSpan', $currentCell->getRowSpan());
                    $rowSpan[$cell] = $currentCell->getRowSpan() - 1;
                } elseif (isset($rowSpan[$cell]) && $rowSpan[$cell] > 0) {
                    $rowSpan[$cell]--;
                    $objWriter->writeAttribute('vMerge', '1');
                }
                // a:txBody
                $objWriter->startElement('a:txBody');
                // a:txBody/a:bodyPr
                $objWriter->startElement('a:bodyPr');
                $objWriter->writeAttribute('wrap', 'square');
                $objWriter->writeAttribute('rtlCol', '0');
                // a:txBody/a:bodyPr/a:spAutoFit
                $objWriter->writeElement('a:spAutoFit', null);
                // a:txBody/a:bodyPr/
                $objWriter->endElement();
                // a:lstStyle
                $objWriter->writeElement('a:lstStyle', null);
                // Write paragraphs
                $this->writeParagraphs($objWriter, $currentCell->getParagraphs());
                $objWriter->endElement();
                // a:tcPr
                $objWriter->startElement('a:tcPr');
                $firstParagraph = $currentCell->getParagraph(0);
                $firstParagraphAlignment = $firstParagraph->getAlignment();

                // Text Direction
                $textDirection = $firstParagraphAlignment->getTextDirection();
                if ($textDirection != Alignment::TEXT_DIRECTION_HORIZONTAL) {
                    $objWriter->writeAttribute('vert', $textDirection);
                }
                // Alignment (horizontal)
                $verticalAlign = $firstParagraphAlignment->getVertical();
                if ($verticalAlign != Alignment::VERTICAL_BASE && $verticalAlign != Alignment::VERTICAL_AUTO) {
                    $objWriter->writeAttribute('anchor', $verticalAlign);
                }

                // Margins
                $objWriter->writeAttribute('marL', CommonDrawing::pixelsToEmu($firstParagraphAlignment->getMarginLeft()));
                $objWriter->writeAttribute('marR', CommonDrawing::pixelsToEmu($firstParagraphAlignment->getMarginRight()));
                $objWriter->writeAttribute('marT', CommonDrawing::pixelsToEmu($firstParagraphAlignment->getMarginTop()));
                $objWriter->writeAttribute('marB', CommonDrawing::pixelsToEmu($firstParagraphAlignment->getMarginBottom()));

                // Determine borders
                $borderLeft = $currentCell->getBorders()->getLeft();
                $borderRight = $currentCell->getBorders()->getRight();
                $borderTop = $currentCell->getBorders()->getTop();
                $borderBottom = $currentCell->getBorders()->getBottom();
                $borderDiagonalDown = $currentCell->getBorders()->getDiagonalDown();
                $borderDiagonalUp = $currentCell->getBorders()->getDiagonalUp();
                // Fix PowerPoint implementation
                if (!is_null($nextCellRight)
                    && $nextCellRight->getBorders()->getRight()->getHashCode() != $defaultBorder->getHashCode()
                ) {
                    $borderRight = $nextCellRight->getBorders()->getLeft();
                }
                if (!is_null($nextCellBelow)
                    && $nextCellBelow->getBorders()->getBottom()->getHashCode() != $defaultBorder->getHashCode()
                ) {
                    $borderBottom = $nextCellBelow->getBorders()->getTop();
                }
                // Write borders
                $this->writeBorder($objWriter, $borderLeft, 'L');
                $this->writeBorder($objWriter, $borderRight, 'R');
                $this->writeBorder($objWriter, $borderTop, 'T');
                $this->writeBorder($objWriter, $borderBottom, 'B');
                $this->writeBorder($objWriter, $borderDiagonalDown, 'TlToBr');
                $this->writeBorder($objWriter, $borderDiagonalUp, 'BlToTr');
                // Fill
                $this->writeFill($objWriter, $currentCell->getFill());
                $objWriter->endElement();
                $objWriter->endElement();
            }
            $objWriter->endElement();
        }
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Write paragraphs
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\RichText\Paragraph[] $paragraphs
     * @param  bool $bIsPlaceholder
     * @throws \Exception
     */
    protected function writeParagraphs(XMLWriter $objWriter, $paragraphs, $bIsPlaceholder = false)
    {
        // Loop trough paragraphs
        foreach ($paragraphs as $paragraph) {
            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            if (!$bIsPlaceholder) {
                $objWriter->startElement('a:pPr');
                $objWriter->writeAttribute('algn', $paragraph->getAlignment()->getHorizontal());
                $objWriter->writeAttribute('fontAlgn', $paragraph->getAlignment()->getVertical());
                $objWriter->writeAttribute('marL', CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getMarginLeft()));
                $objWriter->writeAttribute('marR', CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getMarginRight()));
                $objWriter->writeAttribute('indent', CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getIndent()));
                $objWriter->writeAttribute('lvl', $paragraph->getAlignment()->getLevel());

                $objWriter->startElement('a:lnSpc');
                $objWriter->startElement('a:spcPct');
                $objWriter->writeAttribute('val', $paragraph->getLineSpacing() . "%");
                $objWriter->endElement();
                $objWriter->endElement();

                // Bullet type specified?
                if ($paragraph->getBulletStyle()->getBulletType() != Bullet::TYPE_NONE) {
                    // Color
                    // a:buClr must be before a:buFont (else PowerPoint crashes at launch)
                    if ($paragraph->getBulletStyle()->getBulletColor() instanceof Color) {
                        $objWriter->startElement('a:buClr');
                        $this->writeColor($objWriter, $paragraph->getBulletStyle()->getBulletColor());
                        $objWriter->endElement();
                    }

                    // a:buFont
                    $objWriter->startElement('a:buFont');
                    $objWriter->writeAttribute('typeface', $paragraph->getBulletStyle()->getBulletFont());
                    $objWriter->endElement();

                    if ($paragraph->getBulletStyle()->getBulletType() == Bullet::TYPE_BULLET) {
                        // a:buChar
                        $objWriter->startElement('a:buChar');
                        $objWriter->writeAttribute('char', $paragraph->getBulletStyle()->getBulletChar());
                        $objWriter->endElement();
                    } elseif ($paragraph->getBulletStyle()->getBulletType() == Bullet::TYPE_NUMERIC) {
                        // a:buAutoNum
                        $objWriter->startElement('a:buAutoNum');
                        $objWriter->writeAttribute('type', $paragraph->getBulletStyle()->getBulletNumericStyle());
                        if ($paragraph->getBulletStyle()->getBulletNumericStartAt() != 1) {
                            $objWriter->writeAttribute('startAt', $paragraph->getBulletStyle()->getBulletNumericStartAt());
                        }
                        $objWriter->endElement();
                    }
                }

                $objWriter->endElement();
            }

            // Loop trough rich text elements
            $elements = $paragraph->getRichTextElements();
            foreach ($elements as $element) {
                if ($element instanceof BreakElement) {
                    // a:br
                    $objWriter->writeElement('a:br', null);
                } elseif ($element instanceof Run || $element instanceof TextElement) {
                    // a:r
                    $objWriter->startElement('a:r');

                    // a:rPr
                    if ($element instanceof Run && !$bIsPlaceholder) {
                        // a:rPr
                        $objWriter->startElement('a:rPr');

                        // Lang
                        $objWriter->writeAttribute('lang', ($element->getLanguage() ? $element->getLanguage() : 'en-US'));

                        $objWriter->writeAttributeIf($element->getFont()->isBold(), 'b', '1');
                        $objWriter->writeAttributeIf($element->getFont()->isItalic(), 'i', '1');
                        $objWriter->writeAttributeIf($element->getFont()->isStrikethrough(), 'strike', 'sngStrike');

                        // Size
                        $objWriter->writeAttribute('sz', ($element->getFont()->getSize() * 100));

                        // Character spacing
                        $objWriter->writeAttribute('spc', $element->getFont()->getCharacterSpacing());

                        // Underline
                        $objWriter->writeAttribute('u', $element->getFont()->getUnderline());

                        // Superscript / subscript
                        $objWriter->writeAttributeIf($element->getFont()->isSuperScript(), 'baseline', '30000');
                        $objWriter->writeAttributeIf($element->getFont()->isSubScript(), 'baseline', '-25000');

                        // Color - a:solidFill
                        $objWriter->startElement('a:solidFill');
                        $this->writeColor($objWriter, $element->getFont()->getColor());
                        $objWriter->endElement();

                        // Font - a:latin
                        $objWriter->startElement('a:latin');
                        $objWriter->writeAttribute('typeface', $element->getFont()->getName());
                        $objWriter->endElement();

                        // a:hlinkClick
                        $this->writeHyperlink($objWriter, $element);

                        $objWriter->endElement();
                    }

                    // t
                    $objWriter->startElement('a:t');
                    $objWriter->writeCData(Text::controlCharacterPHP2OOXML($element->getText()));
                    $objWriter->endElement();

                    $objWriter->endElement();
                }
            }

            $objWriter->endElement();
        }
    }

    /**
     * Write Line Shape
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param \PhpOffice\PhpPresentation\Shape\Line $shape
     * @param  int $shapeId
     */
    protected function writeShapeLine(XMLWriter $objWriter, Line $shape, $shapeId)
    {
        // p:sp
        $objWriter->startElement('p:cxnSp');
        // p:nvSpPr
        $objWriter->startElement('p:nvCxnSpPr');
        // p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', $shapeId);
        $objWriter->writeAttribute('name', '');
        $objWriter->endElement();
        // p:cNvCxnSpPr
        $objWriter->writeElement('p:cNvCxnSpPr', null);
        // p:nvPr
        $objWriter->startElement('p:nvPr');
        if ($shape->isPlaceholder()) {
            $objWriter->startElement('p:ph');
            $objWriter->writeAttribute('type', $shape->getPlaceholder()->getType());
            $objWriter->endElement();
        }
        $objWriter->endElement();
        $objWriter->endElement();
        // p:spPr
        $objWriter->startElement('p:spPr');
        // a:xfrm
        $objWriter->startElement('a:xfrm');
        if ($shape->getWidth() >= 0 && $shape->getHeight() >= 0) {
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
        } elseif ($shape->getWidth() < 0 && $shape->getHeight() < 0) {
            // a:off
            $objWriter->startElement('a:off');
            $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($shape->getOffsetX() + $shape->getWidth()));
            $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($shape->getOffsetY() + $shape->getHeight()));
            $objWriter->endElement();
            // a:ext
            $objWriter->startElement('a:ext');
            $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu(-$shape->getWidth()));
            $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu(-$shape->getHeight()));
            $objWriter->endElement();
        } elseif ($shape->getHeight() < 0) {
            $objWriter->writeAttribute('flipV', 1);
            // a:off
            $objWriter->startElement('a:off');
            $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($shape->getOffsetX()));
            $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($shape->getOffsetY() + $shape->getHeight()));
            $objWriter->endElement();
            // a:ext
            $objWriter->startElement('a:ext');
            $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($shape->getWidth()));
            $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu(-$shape->getHeight()));
            $objWriter->endElement();
        } elseif ($shape->getWidth() < 0) {
            $objWriter->writeAttribute('flipV', 1);
            // a:off
            $objWriter->startElement('a:off');
            $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($shape->getOffsetX() + $shape->getWidth()));
            $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($shape->getOffsetY()));
            $objWriter->endElement();
            // a:ext
            $objWriter->startElement('a:ext');
            $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu(-$shape->getWidth()));
            $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($shape->getHeight()));
            $objWriter->endElement();
        }
        $objWriter->endElement();
        // a:prstGeom
        $objWriter->startElement('a:prstGeom');
        $objWriter->writeAttribute('prst', 'line');

        // a:prstGeom/a:avLst
        $objWriter->writeElement('a:avLst');

        $objWriter->endElement();
        $this->writeBorder($objWriter, $shape->getBorder(), '');
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Write Shadow
     * @param XMLWriter $objWriter
     * @param Shadow $oShadow
     */
    protected function writeShadow(XMLWriter $objWriter, $oShadow)
    {
        if (!($oShadow instanceof Shadow)) {
            return;
        }

        if (!$oShadow->isVisible()) {
            return;
        }

        // a:effectLst
        $objWriter->startElement('a:effectLst');

        // a:outerShdw
        $objWriter->startElement('a:outerShdw');
        $objWriter->writeAttribute('blurRad', CommonDrawing::pixelsToEmu($oShadow->getBlurRadius()));
        $objWriter->writeAttribute('dist', CommonDrawing::pixelsToEmu($oShadow->getDistance()));
        $objWriter->writeAttribute('dir', CommonDrawing::degreesToAngle($oShadow->getDirection()));
        $objWriter->writeAttribute('algn', $oShadow->getAlignment());
        $objWriter->writeAttribute('rotWithShape', '0');

        $this->writeColor($objWriter, $oShadow->getColor(), $oShadow->getAlpha());

        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write hyperlink
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param \PhpOffice\PhpPresentation\AbstractShape|\PhpOffice\PhpPresentation\Shape\RichText\TextElement $shape
     */
    protected function writeHyperlink(XMLWriter $objWriter, $shape)
    {
        if (!$shape->hasHyperlink()) {
            return;
        }
        // a:hlinkClick
        $objWriter->startElement('a:hlinkClick');
        $objWriter->writeAttribute('r:id', $shape->getHyperlink()->relationId);
        $objWriter->writeAttribute('tooltip', $shape->getHyperlink()->getTooltip());
        if ($shape->getHyperlink()->isInternal()) {
            $objWriter->writeAttribute('action', $shape->getHyperlink()->getUrl());
        }
        $objWriter->endElement();
    }

    /**
     * Write Note Slide
     * @param Note $pNote
     * @throws \Exception
     * @return  string
     */
    protected function writeNote(Note $pNote)
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:notes
        $objWriter->startElement('p:notes');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        // p:notes/p:cSld
        $objWriter->startElement('p:cSld');

        // p:notes/p:cSld/p:spTree
        $objWriter->startElement('p:spTree');

        // p:notes/p:cSld/p:spTree/p:nvGrpSpPr
        $objWriter->startElement('p:nvGrpSpPr');

        // p:notes/p:cSld/p:spTree/p:nvGrpSpPr/p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '1');
        $objWriter->writeAttribute('name', '');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:nvGrpSpPr/p:cNvGrpSpPr
        $objWriter->writeElement('p:cNvGrpSpPr', null);

        // p:notes/p:cSld/p:spTree/p:nvGrpSpPr/p:nvPr
        $objWriter->writeElement('p:nvPr', null);

        // p:notes/p:cSld/p:spTree/p:nvGrpSpPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:grpSpPr
        $objWriter->startElement('p:grpSpPr');

        // p:notes/p:cSld/p:spTree/p:grpSpPr/a:xfrm
        $objWriter->startElement('a:xfrm');

        // p:notes/p:cSld/p:spTree/p:grpSpPr/a:xfrm/a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pNote->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pNote->getOffsetY()));
        $objWriter->endElement(); // a:off

        // p:notes/p:cSld/p:spTree/p:grpSpPr/a:xfrm/a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pNote->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pNote->getExtentY()));
        $objWriter->endElement(); // a:ext

        // p:notes/p:cSld/p:spTree/p:grpSpPr/a:xfrm/a:chOff
        $objWriter->startElement('a:chOff');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pNote->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pNote->getOffsetY()));
        $objWriter->endElement(); // a:chOff

        // p:notes/p:cSld/p:spTree/p:grpSpPr/a:xfrm/a:chExt
        $objWriter->startElement('a:chExt');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pNote->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pNote->getExtentY()));
        $objWriter->endElement(); // a:chExt

        // p:notes/p:cSld/p:spTree/p:grpSpPr/a:xfrm
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:grpSpPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]
        $objWriter->startElement('p:sp');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr
        $objWriter->startElement('p:nvSpPr');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr/p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '2');
        $objWriter->writeAttribute('name', 'Slide Image Placeholder 1');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr/p:cNvSpPr
        $objWriter->startElement('p:cNvSpPr');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr/p:cNvSpPr/a:spLocks
        $objWriter->startElement('a:spLocks');
        $objWriter->writeAttribute('noGrp', '1');
        $objWriter->writeAttribute('noRot', '1');
        $objWriter->writeAttribute('noChangeAspect', '1');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr/p:cNvSpPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr/p:nvPr
        $objWriter->startElement('p:nvPr');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr/p:nvPr/p:ph
        $objWriter->startElement('p:ph');
        $objWriter->writeAttribute('type', 'sldImg');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr/p:nvPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:nvSpPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr
        $objWriter->startElement('p:spPr');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm
        $objWriter->startElement('a:xfrm');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', 0);
        $objWriter->writeAttribute('y', 0);
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm/a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu(round($pNote->getExtentX() / 2)));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu(round($pNote->getExtentY() / 2)));
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:xfrm
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:prstGeom
        $objWriter->startElement('a:prstGeom');
        $objWriter->writeAttribute('prst', 'rect');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:prstGeom/a:avLst
        $objWriter->writeElement('a:avLst', null);

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:prstGeom
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:noFill
        $objWriter->writeElement('a:noFill', null);

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:ln
        $objWriter->startElement('a:ln');
        $objWriter->writeAttribute('w', '12700');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:ln/a:solidFill
        $objWriter->startElement('a:solidFill');

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:ln/a:solidFill/a:prstClr
        $objWriter->startElement('a:prstClr');
        $objWriter->writeAttribute('val', 'black');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:ln/a:solidFill
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr/a:ln
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]/p:spPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[1]
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]
        $objWriter->startElement('p:sp');

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr
        $objWriter->startElement('p:nvSpPr');

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '3');
        $objWriter->writeAttribute('name', 'Notes Placeholder');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:cNvSpPr
        $objWriter->startElement('p:cNvSpPr');

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:cNvSpPr/a:spLocks
        $objWriter->startElement('a:spLocks');
        $objWriter->writeAttribute('noGrp', '1');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:cNvSpPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:nvPr
        $objWriter->startElement('p:nvPr');

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:nvPr/p:ph
        $objWriter->startElement('p:ph');
        $objWriter->writeAttribute('type', 'body');
        $objWriter->writeAttribute('idx', '1');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr/p:nvPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:nvSpPr
        $objWriter->endElement();

        // START notes print below rectangle section
        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr
        $objWriter->startElement('p:spPr');

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm
        $objWriter->startElement('a:xfrm');

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm/a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pNote->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu(round($pNote->getExtentY() / 2) + $pNote->getOffsetY()));
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm/a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', '5486400');
        $objWriter->writeAttribute('cy', '3600450');
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:xfrm
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:prstGeom
        $objWriter->startElement('a:prstGeom');
        $objWriter->writeAttribute('prst', 'rect');

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:prstGeom/a:avLst
        $objWriter->writeElement('a:avLst', null);

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr/a:prstGeom
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:spPr
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:txBody
        $objWriter->startElement('p:txBody');

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:txBody/a:bodyPr
        $objWriter->writeElement('a:bodyPr', null);
        // p:notes/p:cSld/p:spTree/p:sp[2]/p:txBody/a:lstStyle
        $objWriter->writeElement('a:lstStyle', null);

        // Loop shapes
        $shapes = $pNote->getShapeCollection();
        foreach ($shapes as $shape) {
            // Check type
            if ($shape instanceof RichText) {
                $paragraphs = $shape->getParagraphs();
                $this->writeParagraphs($objWriter, $paragraphs);
            }
        }

        // p:notes/p:cSld/p:spTree/p:sp[2]/p:txBody
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree/p:sp[2]
        $objWriter->endElement();

        // p:notes/p:cSld/p:spTree
        $objWriter->endElement();

        // p:notes/p:cSld
        $objWriter->endElement();

        // p:notes
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write chart
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param \PhpOffice\PhpPresentation\Shape\Chart $shape
     * @param  int $shapeId
     */
    protected function writeShapeChart(XMLWriter $objWriter, ShapeChart $shape, $shapeId)
    {
        // p:graphicFrame
        $objWriter->startElement('p:graphicFrame');
        // p:nvGraphicFramePr
        $objWriter->startElement('p:nvGraphicFramePr');
        // p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', $shapeId);
        $objWriter->writeAttribute('name', $shape->getName());
        $objWriter->writeAttribute('descr', $shape->getDescription());
        $objWriter->endElement();
        // p:cNvGraphicFramePr
        $objWriter->writeElement('p:cNvGraphicFramePr', null);
        // p:nvPr
        $objWriter->startElement('p:nvPr');
        if ($shape->isPlaceholder()) {
            $objWriter->startElement('p:ph');
            $objWriter->writeAttribute('type', $shape->getPlaceholder()->getType());
            $objWriter->endElement();
        }
        $objWriter->endElement();
        $objWriter->endElement();
        // p:xfrm
        $objWriter->startElement('p:xfrm');
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
        // a:graphic
        $objWriter->startElement('a:graphic');
        // a:graphicData
        $objWriter->startElement('a:graphicData');
        $objWriter->writeAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
        // c:chart
        $objWriter->startElement('c:chart');
        $objWriter->writeAttribute('xmlns:c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('r:id', $shape->relationId);
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Write pic
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\AbstractGraphic $shape
     * @param  int $shapeId
     * @throws \Exception
     */
    protected function writeShapePic(XMLWriter $objWriter, AbstractGraphic $shape, $shapeId)
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
            $objWriter->writeAttribute('r:embed', $shape->relationId);
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
        if ($shape->getBorder()->getLineStyle() != Border::LINE_NONE) {
            $this->writeBorder($objWriter, $shape->getBorder(), '');
        }
        if ($shape->getShadow()->isVisible()) {
            $this->writeShadow($objWriter, $shape->getShadow());
        }
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Write group
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param \PhpOffice\PhpPresentation\Shape\Group $group
     * @param  int $shapeId
     */
    protected function writeShapeGroup(XMLWriter $objWriter, Group $group, &$shapeId)
    {
        // p:grpSp
        $objWriter->startElement('p:grpSp');
        // p:nvGrpSpPr
        $objWriter->startElement('p:nvGrpSpPr');
        // p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('name', 'Group ' . $shapeId++);
        $objWriter->writeAttribute('id', $shapeId);
        $objWriter->endElement(); // p:cNvPr
        // NOTE: Re: $shapeId This seems to be how PowerPoint 2010 does business.
        // p:cNvGrpSpPr
        $objWriter->writeElement('p:cNvGrpSpPr', null);
        // p:nvPr
        $objWriter->writeElement('p:nvPr', null);
        $objWriter->endElement(); // p:nvGrpSpPr
        // p:grpSpPr
        $objWriter->startElement('p:grpSpPr');
        // a:xfrm
        $objWriter->startElement('a:xfrm');
        // a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($group->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($group->getOffsetY()));
        $objWriter->endElement(); // a:off
        // a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($group->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($group->getExtentY()));
        $objWriter->endElement(); // a:ext
        // a:chOff
        $objWriter->startElement('a:chOff');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($group->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($group->getOffsetY()));
        $objWriter->endElement(); // a:chOff
        // a:chExt
        $objWriter->startElement('a:chExt');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($group->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($group->getExtentY()));
        $objWriter->endElement(); // a:chExt
        $objWriter->endElement(); // a:xfrm
        $objWriter->endElement(); // p:grpSpPr

        $this->writeShapeCollection($objWriter, $group->getShapeCollection(), $shapeId);

        $objWriter->endElement(); // p:grpSp
    }

    /**
     * @param \PhpOffice\PhpPresentation\Slide\AbstractSlide $pSlide
     * @param $objWriter
     */
    protected function writeSlideBackground(AbstractSlideAlias $pSlide, XMLWriter $objWriter)
    {
        if (!($pSlide->getBackground() instanceof Slide\AbstractBackground)) {
            return;
        }
        $oBackground = $pSlide->getBackground();
        // p:bg
        $objWriter->startElement('p:bg');
        if ($oBackground instanceof Slide\Background\Color) {
            // p:bgPr
            $objWriter->startElement('p:bgPr');
            // a:solidFill
            $objWriter->startElement('a:solidFill');
            // a:srgbClr
            $objWriter->startElement('a:srgbClr');
            $objWriter->writeAttribute('val', $oBackground->getColor()->getRGB());
            $objWriter->endElement();
            // > a:solidFill
            $objWriter->endElement();

            // p:bgPr/a:effectLst
            $objWriter->writeElement('a:effectLst');

            // > p:bgPr
            $objWriter->endElement();
        }
        if ($oBackground instanceof Slide\Background\Image) {
            // p:bgPr
            $objWriter->startElement('p:bgPr');
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
            // > p:bgPr
            $objWriter->endElement();
        }
        /**
         * @link : http://www.officeopenxml.com/prSlide-background.php
         */
        if ($oBackground instanceof Slide\Background\SchemeColor) {
            // p:bgRef
            $objWriter->startElement('p:bgRef');
            $objWriter->writeAttribute('idx', '1001');
            // a:schemeClr
            $objWriter->startElement('a:schemeClr');
            $objWriter->writeAttribute('val', $oBackground->getSchemeColor()->getValue());
            $objWriter->endElement();
            // > p:bgRef
            $objWriter->endElement();
        }
        // > p:bg
        $objWriter->endElement();
    }


    /**
     * Write Transition Slide
     * @link http://officeopenxml.com/prSlide-transitions.php
     * @param XMLWriter $objWriter
     * @param Slide\Transition $transition
     */
    protected function writeSlideTransition(XMLWriter $objWriter, $transition)
    {
        if (!$transition instanceof Slide\Transition) {
            return;
        }
        $objWriter->startElement('p:transition');
        if (!is_null($transition->getSpeed())) {
            $objWriter->writeAttribute('spd', $transition->getSpeed());
        }
        $objWriter->writeAttribute('advClick', $transition->hasManualTrigger() ? '1' : '0');
        if ($transition->hasTimeTrigger()) {
            $objWriter->writeAttribute('advTm', $transition->getAdvanceTimeTrigger());
        }

        switch ($transition->getTransitionType()) {
            case Slide\Transition::TRANSITION_BLINDS_HORIZONTAL:
                $objWriter->startElement('p:blinds');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_BLINDS_VERTICAL:
                $objWriter->startElement('p:blinds');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_CHECKER_HORIZONTAL:
                $objWriter->startElement('p:checker');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_CHECKER_VERTICAL:
                $objWriter->startElement('p:checker');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_CIRCLE_HORIZONTAL:
                $objWriter->startElement('p:circle');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_CIRCLE_VERTICAL:
                $objWriter->startElement('p:circle');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COMB_HORIZONTAL:
                $objWriter->startElement('p:comb');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COMB_VERTICAL:
                $objWriter->startElement('p:comb');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COVER_DOWN:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'd');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COVER_LEFT:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'l');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COVER_LEFT_DOWN:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'ld');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COVER_LEFT_UP:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'lu');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COVER_RIGHT:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'r');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COVER_RIGHT_DOWN:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'rd');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COVER_RIGHT_UP:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'ru');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_COVER_UP:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'u');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_CUT:
                $objWriter->writeElement('p:cut');
                break;
            case Slide\Transition::TRANSITION_DIAMOND:
                $objWriter->writeElement('p:diamond');
                break;
            case Slide\Transition::TRANSITION_DISSOLVE:
                $objWriter->writeElement('p:dissolve');
                break;
            case Slide\Transition::TRANSITION_FADE:
                $objWriter->writeElement('p:fade');
                break;
            case Slide\Transition::TRANSITION_NEWSFLASH:
                $objWriter->writeElement('p:newsflash');
                break;
            case Slide\Transition::TRANSITION_PLUS:
                $objWriter->writeElement('p:plus');
                break;
            case Slide\Transition::TRANSITION_PULL_DOWN:
                $objWriter->startElement('p:pull');
                $objWriter->writeAttribute('dir', 'd');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_PULL_LEFT:
                $objWriter->startElement('p:pull');
                $objWriter->writeAttribute('dir', 'l');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_PULL_RIGHT:
                $objWriter->startElement('p:pull');
                $objWriter->writeAttribute('dir', 'r');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_PULL_UP:
                $objWriter->startElement('p:pull');
                $objWriter->writeAttribute('dir', 'u');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_PUSH_DOWN:
                $objWriter->startElement('p:push');
                $objWriter->writeAttribute('dir', 'd');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_PUSH_LEFT:
                $objWriter->startElement('p:push');
                $objWriter->writeAttribute('dir', 'l');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_PUSH_RIGHT:
                $objWriter->startElement('p:push');
                $objWriter->writeAttribute('dir', 'r');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_PUSH_UP:
                $objWriter->startElement('p:push');
                $objWriter->writeAttribute('dir', 'u');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_RANDOM:
                $objWriter->writeElement('p:random');
                break;
            case Slide\Transition::TRANSITION_RANDOMBAR_HORIZONTAL:
                $objWriter->startElement('p:randomBar');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_RANDOMBAR_VERTICAL:
                $objWriter->startElement('p:randomBar');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_SPLIT_IN_HORIZONTAL:
                $objWriter->startElement('p:split');
                $objWriter->writeAttribute('dir', 'in');
                $objWriter->writeAttribute('orient', 'horz');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_SPLIT_OUT_HORIZONTAL:
                $objWriter->startElement('p:split');
                $objWriter->writeAttribute('dir', 'out');
                $objWriter->writeAttribute('orient', 'horz');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_SPLIT_IN_VERTICAL:
                $objWriter->startElement('p:split');
                $objWriter->writeAttribute('dir', 'in');
                $objWriter->writeAttribute('orient', 'vert');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_SPLIT_OUT_VERTICAL:
                $objWriter->startElement('p:split');
                $objWriter->writeAttribute('dir', 'out');
                $objWriter->writeAttribute('orient', 'vert');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_STRIPS_LEFT_DOWN:
                $objWriter->startElement('p:strips');
                $objWriter->writeAttribute('dir', 'ld');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_STRIPS_LEFT_UP:
                $objWriter->startElement('p:strips');
                $objWriter->writeAttribute('dir', 'lu');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_STRIPS_RIGHT_DOWN:
                $objWriter->startElement('p:strips');
                $objWriter->writeAttribute('dir', 'rd');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_STRIPS_RIGHT_UP:
                $objWriter->startElement('p:strips');
                $objWriter->writeAttribute('dir', 'ru');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_WEDGE:
                $objWriter->writeElement('p:wedge');
                break;
            case Slide\Transition::TRANSITION_WIPE_DOWN:
                $objWriter->startElement('p:wipe');
                $objWriter->writeAttribute('dir', 'd');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_WIPE_LEFT:
                $objWriter->startElement('p:wipe');
                $objWriter->writeAttribute('dir', 'l');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_WIPE_RIGHT:
                $objWriter->startElement('p:wipe');
                $objWriter->writeAttribute('dir', 'r');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_WIPE_UP:
                $objWriter->startElement('p:wipe');
                $objWriter->writeAttribute('dir', 'u');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_ZOOM_IN:
                $objWriter->startElement('p:zoom');
                $objWriter->writeAttribute('dir', 'in');
                $objWriter->endElement();
                break;
            case Slide\Transition::TRANSITION_ZOOM_OUT:
                $objWriter->startElement('p:zoom');
                $objWriter->writeAttribute('dir', 'out');
                $objWriter->endElement();
                break;
        }

        $objWriter->endElement();
    }

    private function getGUID()
    {
        if (function_exists('com_create_guid')) {
            return com_create_guid();
        } else {
            mt_srand((double)microtime() * 10000);//optional for php 4.2.0 and up.
            $charid = strtoupper(md5(uniqid(rand(), true)));
            $hyphen = chr(45);// "-"
            $uuid = chr(123)// "{"
                . substr($charid, 0, 8) . $hyphen
                . substr($charid, 8, 4) . $hyphen
                . substr($charid, 12, 4) . $hyphen
                . substr($charid, 16, 4) . $hyphen
                . substr($charid, 20, 12)
                . chr(125);// "}"
            return $uuid;
        }
    }
}
