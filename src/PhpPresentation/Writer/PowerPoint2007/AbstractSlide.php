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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use ArrayObject;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\Exception\UndefinedChartTypeException;
use PhpOffice\PhpPresentation\Shape\AbstractGraphic;
use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Shape\Chart as ShapeChart;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Drawing\File as ShapeDrawingFile;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd as ShapeDrawingGd;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table as ShapeTable;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\AbstractSlide as AbstractSlideAlias;
use PhpOffice\PhpPresentation\Slide\Note;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Shadow;

abstract class AbstractSlide extends AbstractDecoratorWriter
{
    /**
     * @return mixed
     */
    protected function writeDrawingRelations(AbstractSlideAlias $pSlideMaster, XMLWriter $objWriter, int $relId)
    {
        if (count($pSlideMaster->getShapeCollection()) > 0) {
            // Loop trough images and write relationships
            foreach ($pSlideMaster->getShapeCollection() as $shape) {
                if ($shape instanceof ShapeDrawingFile || $shape instanceof ShapeDrawingGd) {
                    // Write relationship for image drawing
                    $this->writeRelationship(
                        $objWriter,
                        $relId,
                        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                        '../media/' . str_replace(' ', '_', $shape->getIndexedFilename())
                    );
                    $shape->relationId = 'rId' . $relId;
                    ++$relId;
                } elseif ($shape instanceof ShapeChart) {
                    // Write relationship for chart drawing
                    $this->writeRelationship(
                        $objWriter,
                        $relId,
                        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                        '../charts/' . $shape->getIndexedFilename()
                    );
                    $shape->relationId = 'rId' . $relId;
                    ++$relId;
                } elseif ($shape instanceof Group) {
                    foreach ($shape->getShapeCollection() as $subShape) {
                        if ($subShape instanceof ShapeDrawingFile ||
                            $subShape instanceof ShapeDrawingGd
                        ) {
                            // Write relationship for image drawing
                            $this->writeRelationship(
                                $objWriter,
                                $relId,
                                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                                '../media/' . str_replace(' ', '_', $subShape->getIndexedFilename())
                            );
                            $subShape->relationId = 'rId' . $relId;
                            ++$relId;
                        } elseif ($subShape instanceof ShapeChart) {
                            // Write relationship for chart drawing
                            $this->writeRelationship(
                                $objWriter,
                                $relId,
                                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
                                '../charts/' . $subShape->getIndexedFilename()
                            );
                            $subShape->relationId = 'rId' . $relId;
                            ++$relId;
                        }
                    }
                }
            }
        }

        return $relId;
    }

    /**
     * Note : $shapeId needs to start to 1
     *  The animation is applied to the shape which is next to the target shape.
     *
     * @param array<int, AbstractShape>|ArrayObject<int, AbstractShape> $shapes
     * @param int $shapeId
     */
    protected function writeShapeCollection(XMLWriter $objWriter, $shapes = [], &$shapeId = 1): void
    {
        if (0 == count($shapes)) {
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
            } elseif ($shape instanceof AutoShape) {
                $this->writeShapeAutoShape($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Group) {
                $this->writeShapeGroup($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Comment) {
            } else {
                throw new UndefinedChartTypeException();
            }
        }
    }

    /**
     * Write txt.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeShapeText(XMLWriter $objWriter, RichText $shape, int $shapeId): void
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
            $objWriter->writeAttribute('name', $shape->getName());
        }
        // Hyperlink
        if ($shape->hasHyperlink()) {
            $this->writeHyperlink($objWriter, $shape);
        }
        // > p:sp\p:nvSpPr
        $objWriter->endElement();
        // p:sp\p:nvSpPr\p:cNvSpPr
        $objWriter->startElement('p:cNvSpPr');
        $objWriter->writeAttribute('txBox', '1');
        $objWriter->endElement();
        // p:sp\p:nvSpPr\p:nvPr
        if ($shape->isPlaceholder()) {
            $objWriter->startElement('p:nvPr');
            $objWriter->startElement('p:ph');
            $objWriter->writeAttribute('type', $shape->getPlaceholder()->getType());
            if (null !== $shape->getPlaceholder()->getIdx()) {
                $objWriter->writeAttribute('idx', $shape->getPlaceholder()->getIdx());
            }
            $objWriter->endElement();
            $objWriter->endElement();
        } else {
            $objWriter->writeElement('p:nvPr', null);
        }
        // > p:sp\p:nvSpPr
        $objWriter->endElement();
        // p:sp\p:spPr
        $objWriter->startElement('p:spPr');

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
            // Vertical alignment
            $verticalAlign = $shape->getActiveParagraph()->getAlignment()->getVertical();
            if (Alignment::VERTICAL_BASE != $verticalAlign && Alignment::VERTICAL_AUTO != $verticalAlign) {
                $objWriter->writeAttribute('anchor', $verticalAlign);
            }
            $objWriter->writeAttribute('anchorCtr', $shape->getVerticalAlignCenter());
            if (RichText::WRAP_SQUARE != $shape->getWrap()) {
                $objWriter->writeAttribute('wrap', $shape->getWrap());
            }
            $objWriter->writeAttribute('rtlCol', '0');
            if (RichText::OVERFLOW_OVERFLOW != $shape->getHorizontalOverflow()) {
                $objWriter->writeAttribute('horzOverflow', $shape->getHorizontalOverflow());
            }
            if (RichText::OVERFLOW_OVERFLOW != $shape->getVerticalOverflow()) {
                $objWriter->writeAttribute('vertOverflow', $shape->getVerticalOverflow());
            }
            if ($shape->isUpright()) {
                $objWriter->writeAttribute('upright', '1');
            }
            $objWriter->writeAttribute('vert', $shape->isVertical() ? 'vert' : 'horz');
            $objWriter->writeAttribute('bIns', CommonDrawing::pixelsToEmu($shape->getInsetBottom()));
            $objWriter->writeAttribute('lIns', CommonDrawing::pixelsToEmu($shape->getInsetLeft()));
            $objWriter->writeAttribute('rIns', CommonDrawing::pixelsToEmu($shape->getInsetRight()));
            $objWriter->writeAttribute('tIns', CommonDrawing::pixelsToEmu($shape->getInsetTop()));
            if ($shape->getColumns() != 1) {
                $objWriter->writeAttribute('numCol', $shape->getColumns());
                $objWriter->writeAttribute('spcCol', CommonDrawing::pixelsToEmu($shape->getColumnSpacing()));
            }
            // a:spAutoFit
            $objWriter->startElement('a:' . $shape->getAutoFit());
            if (RichText::AUTOFIT_NORMAL == $shape->getAutoFit()) {
                if (null !== $shape->getFontScale()) {
                    $objWriter->writeAttribute('fontScale', $shape->getFontScale() * 1000);
                }
                if (null !== $shape->getLineSpaceReduction()) {
                    $objWriter->writeAttribute('lnSpcReduction', $shape->getLineSpaceReduction() * 1000);
                }
            }
            $objWriter->endElement();
        }
        $objWriter->endElement();
        // a:lstStyle
        $objWriter->writeElement('a:lstStyle', null);
        if ($shape->isPlaceholder() &&
            (Placeholder::PH_TYPE_SLIDENUM == $shape->getPlaceholder()->getType() ||
                Placeholder::PH_TYPE_DATETIME == $shape->getPlaceholder()->getType())
        ) {
            $objWriter->startElement('a:p');

            // Paragraph Style
            $paragraphs = $shape->getParagraphs();
            if (!empty($paragraphs)) {
                $paragraph = &$paragraphs[0];
                $this->writeParagraphStyles($objWriter, $paragraph, true);
            }

            $objWriter->startElement('a:fld');
            $objWriter->writeAttribute('id', $this->getGUID());
            $objWriter->writeAttribute('type', (
                Placeholder::PH_TYPE_SLIDENUM == $shape->getPlaceholder()->getType() ? 'slidenum' : 'datetime'
            ));

            if (isset($paragraph)) {
                $elements = $paragraph->getRichTextElements();
                if (!empty($elements)) {
                    $element = &$elements[0];
                    if ($element instanceof Run) {
                        $this->writeRunStyles($objWriter, $element);
                    }
                }
            }

            $objWriter->writeElement('a:t', (
                Placeholder::PH_TYPE_SLIDENUM == $shape->getPlaceholder()->getType() ? '<nr.>' : '03-04-05'
            ));
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
     * Write table.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeShapeTable(XMLWriter $objWriter, ShapeTable $shape, int $shapeId): void
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
        for ($cell = 0; $cell < $countCells; ++$cell) {
            //  p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tblGrid/a:gridCol
            $objWriter->startElement('a:gridCol');
            // Calculate column width
            $width = $shape->getRow(0)->getCell($cell)->getWidth();
            if (0 == $width) {
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
        $colSpan = $rowSpan = [];
        // Default border style
        $defaultBorder = new Border();
        // Write rows
        $countRows = count($shape->getRows());
        for ($row = 0; $row < $countRows; ++$row) {
            // p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr
            $objWriter->startElement('a:tr');
            $objWriter->writeAttribute('h', CommonDrawing::pixelsToEmu($shape->getRow($row)->getHeight()));
            // Write cells
            $countCells = count($shape->getRow($row)->getCells());
            for ($cell = 0; $cell < $countCells; ++$cell) {
                // Current cell
                $currentCell = $shape->getRow($row)->getCell($cell);
                // Next cell right
                $hasNextCellRight = $shape->getRow($row)->hasCell($cell + 1);
                // Next cell below
                $hasNextRowBelow = $shape->hasRow($row + 1);
                // a:tc
                $objWriter->startElement('a:tc');
                // Colspan
                if ($currentCell->getColSpan() > 1) {
                    $objWriter->writeAttribute('gridSpan', $currentCell->getColSpan());
                    $colSpan[$row] = $currentCell->getColSpan() - 1;
                } elseif (isset($colSpan[$row]) && $colSpan[$row] > 0) {
                    --$colSpan[$row];
                    $objWriter->writeAttribute('hMerge', '1');
                }
                // Rowspan
                if ($currentCell->getRowSpan() > 1) {
                    $objWriter->writeAttribute('rowSpan', $currentCell->getRowSpan());
                    $rowSpan[$cell] = $currentCell->getRowSpan() - 1;
                } elseif (isset($rowSpan[$cell]) && $rowSpan[$cell] > 0) {
                    --$rowSpan[$cell];
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
                if (Alignment::TEXT_DIRECTION_HORIZONTAL != $textDirection) {
                    $objWriter->writeAttribute('vert', $textDirection);
                }
                // Alignment (horizontal)
                $verticalAlign = $firstParagraphAlignment->getVertical();
                if (Alignment::VERTICAL_BASE != $verticalAlign && Alignment::VERTICAL_AUTO != $verticalAlign) {
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
                if ($hasNextCellRight) {
                    $nextCellRight = $shape->getRow($row)->getCell($cell + 1);
                    if ($nextCellRight->getBorders()->getRight()->getHashCode() != $defaultBorder->getHashCode()) {
                        $borderRight = $nextCellRight->getBorders()->getLeft();
                    }
                }
                if ($hasNextRowBelow) {
                    $nextRowBelow = $shape->getRow($row + 1);
                    $nextCellBelow = $nextRowBelow->getCell($cell);
                    if ($nextCellBelow->getBorders()->getBottom()->getHashCode() != $defaultBorder->getHashCode()) {
                        $borderBottom = $nextCellBelow->getBorders()->getTop();
                    }
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
     * Write paragraphs.
     *
     * @param XMLWriter $objWriter XML Writer
     * @param array<Paragraph> $paragraphs
     */
    protected function writeParagraphs(XMLWriter $objWriter, array $paragraphs): void
    {
        // Loop trough paragraphs
        foreach ($paragraphs as $paragraph) {
            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $this->writeParagraphStyles($objWriter, $paragraph, false);

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
                    if ($element instanceof Run) {
                        $this->writeRunStyles($objWriter, $element);
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
     * Write Paragraph Styles (a:pPr).
     */
    protected function writeParagraphStyles(XMLWriter $objWriter, Paragraph $paragraph, bool $isPlaceholder = false): void
    {
        $objWriter->startElement('a:pPr');
        $objWriter->writeAttribute('algn', $paragraph->getAlignment()->getHorizontal());
        $objWriter->writeAttribute('rtl', $paragraph->getAlignment()->isRTL() ? '1' : '0');
        $objWriter->writeAttribute('fontAlgn', $paragraph->getAlignment()->getVertical());
        $objWriter->writeAttribute('marL', CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getMarginLeft()));
        $objWriter->writeAttribute('marR', CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getMarginRight()));
        $objWriter->writeAttribute('indent', (int) CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getIndent()));
        $objWriter->writeAttribute('lvl', $paragraph->getAlignment()->getLevel());

        $objWriter->startElement('a:lnSpc');
        if ($paragraph->getLineSpacingMode() == Paragraph::LINE_SPACING_MODE_POINT) {
            $objWriter->startElement('a:spcPts');
            $objWriter->writeAttribute('val', $paragraph->getLineSpacing() * 100);
            $objWriter->endElement();
        } else {
            $objWriter->startElement('a:spcPct');
            $objWriter->writeAttribute('val', $paragraph->getLineSpacing() * 1000);
            $objWriter->endElement();
        }
        $objWriter->endElement();

        $objWriter->startElement('a:spcBef');
        $objWriter->startElement('a:spcPts');
        $objWriter->writeAttribute('val', $paragraph->getSpacingBefore() * 100);
        $objWriter->endElement();
        $objWriter->endElement();

        $objWriter->startElement('a:spcAft');
        $objWriter->startElement('a:spcPts');
        $objWriter->writeAttribute('val', $paragraph->getSpacingAfter() * 100);
        $objWriter->endElement();
        $objWriter->endElement();

        if (!$isPlaceholder) {
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
        }

        $objWriter->endElement();
    }

    /**
     * Write RichTextElement Styles (a:pPr).
     */
    protected function writeRunStyles(XMLWriter $objWriter, Run $element): void
    {
        // a:rPr
        $objWriter->startElement('a:rPr');

        // Lang
        $objWriter->writeAttribute('lang', ($element->getLanguage() ? $element->getLanguage() : 'en-US'));

        $objWriter->writeAttributeIf($element->getFont()->isBold(), 'b', '1');
        $objWriter->writeAttributeIf($element->getFont()->isItalic(), 'i', '1');

        // Strikethrough
        $objWriter->writeAttribute('strike', $element->getFont()->getStrikethrough());

        // Size
        $objWriter->writeAttribute('sz', ($element->getFont()->getSize() * 100));

        // Character spacing
        $objWriter->writeAttribute('spc', $element->getFont()->getCharacterSpacing());

        // Underline
        $objWriter->writeAttribute('u', $element->getFont()->getUnderline());

        // Capitalization
        $objWriter->writeAttribute('cap', $element->getFont()->getCapitalization());

        // Baseline
        $objWriter->writeAttributeIf($element->getFont()->getBaseline() !== 0, 'baseline', $element->getFont()->getBaseline());

        // Color - a:solidFill
        $objWriter->startElement('a:solidFill');
        $this->writeColor($objWriter, $element->getFont()->getColor());
        $objWriter->endElement();

        // Font
        // - a:latin
        // - a:ea
        // - a:cs
        $objWriter->startElement('a:' . $element->getFont()->getFormat());
        $objWriter->writeAttribute('typeface', $element->getFont()->getName());
        if ($element->getFont()->getPanose() !== '') {
            $panose = array_map(function (string $value) {
                return '0' . $value;
            }, str_split($element->getFont()->getPanose()));

            $objWriter->writeAttribute('panose', implode('', $panose));
        }
        $objWriter->writeAttributeIf(
            $element->getFont()->getPitchFamily() !== 0,
            'pitchFamily',
            $element->getFont()->getPitchFamily()
        );
        $objWriter->writeAttributeIf(
            $element->getFont()->getCharset() !== Font::CHARSET_DEFAULT,
            'charset',
            dechex($element->getFont()->getCharset())
        );
        $objWriter->endElement();

        // a:hlinkClick
        $this->writeHyperlink($objWriter, $element);

        $objWriter->endElement();
    }

    /**
     * Write Line Shape.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeShapeLine(XMLWriter $objWriter, Line $shape, int $shapeId): void
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
     * Write Shadow.
     */
    protected function writeShadow(XMLWriter $objWriter, Shadow $oShadow): void
    {
        if (!$oShadow->isVisible()) {
            return;
        }

        // a:effectLst
        $objWriter->startElement('a:effectLst');

        // a:outerShdw
        $objWriter->startElement('a:' . $oShadow->getType());
        $objWriter->writeAttribute('blurRad', CommonDrawing::pixelsToEmu($oShadow->getBlurRadius()));
        $objWriter->writeAttribute('dist', CommonDrawing::pixelsToEmu($oShadow->getDistance()));
        $objWriter->writeAttribute('dir', CommonDrawing::degreesToAngle((int) $oShadow->getDirection()));
        $objWriter->writeAttribute('algn', $oShadow->getAlignment());
        $objWriter->writeAttribute('rotWithShape', '0');

        $this->writeColor($objWriter, $oShadow->getColor(), $oShadow->getAlpha());

        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write hyperlink.
     *
     * @param XMLWriter $objWriter XML Writer
     * @param AbstractShape|TextElement $shape
     */
    protected function writeHyperlink(XMLWriter $objWriter, $shape): void
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

        if ($shape->getHyperlink()->isTextColorUsed()) {
            $objWriter->startElement('a:extLst');
            $objWriter->startElement('a:ext');
            $objWriter->writeAttribute('uri', '{A12FA001-AC4F-418D-AE19-62706E023703}');
            $objWriter->startElement('ahyp:hlinkClr');
            $objWriter->writeAttribute('xmlns:ahyp', 'http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor');
            $objWriter->writeAttribute('val', 'tx');
            $objWriter->endElement();
            $objWriter->endElement();
            $objWriter->endElement();
        }

        $objWriter->endElement();
    }

    /**
     * Write Note Slide.
     */
    protected function writeNote(Note $pNote): string
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
     * Write AutoShape.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeShapeAutoShape(XMLWriter $objWriter, AutoShape $shape, int $shapeId): void
    {
        // p:sp
        $objWriter->startElement('p:sp');

        // p:sp\p:nvSpPr
        $objWriter->startElement('p:nvSpPr');
        // p:sp\p:nvSpPr\p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', $shapeId);
        $objWriter->writeAttribute('name', '');
        $objWriter->writeAttribute('descr', '');
        // p:sp\p:nvSpPr\p:cNvPr\
        $objWriter->endElement();
        // p:sp\p:nvSpPr\p:cNvSpPr
        $objWriter->writeElement('p:cNvSpPr');
        // p:sp\p:nvSpPr\p:nvPr
        $objWriter->writeElement('p:nvPr');
        // p:sp\p:nvSpPr\
        $objWriter->endElement();

        // p:sp\p:spPr
        $objWriter->startElement('p:spPr');

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
        // p:sp\p:spPr\a:xfrm\
        $objWriter->endElement();

        // p:sp\p:spPr\a:prstGeom
        $objWriter->startElement('a:prstGeom');
        $objWriter->writeAttribute('prst', $shape->getType());
        // p:sp\p:spPr\a:prstGeom\a:avLst
        $objWriter->writeElement('a:avLst');
        // p:sp\p:spPr\a:prstGeom\
        $objWriter->endElement();
        // Fill
        $this->writeFill($objWriter, $shape->getFill());
        // Outline
        $this->writeOutline($objWriter, $shape->getOutline());

        // p:sp\p:spPr\
        $objWriter->endElement();
        // p:sp\p:txBody
        $objWriter->startElement('p:txBody');
        // p:sp\p:txBody\a:bodyPr
        $objWriter->startElement('a:bodyPr');
        $objWriter->writeAttribute('vertOverflow', 'clip');
        $objWriter->writeAttribute('rtlCol', '0');
        $objWriter->writeAttribute('anchor', 'ctr');
        // p:sp\p:txBody\a:bodyPr\
        $objWriter->endElement();

        // p:sp\p:txBody\a:lstStyle
        $objWriter->writeElement('a:lstStyle');

        // p:sp\p:txBody\a:p
        $objWriter->startElement('a:p');

        // p:sp\p:txBody\a:p\a:pPr
        $objWriter->writeElementBlock('a:pPr', [
            'algn' => 'ctr',
        ]);
        // p:sp\p:txBody\a:p\a:r
        $objWriter->startElement('a:r');
        // p:sp\p:txBody\a:p\a:r\a:t
        $objWriter->startElement('a:t');
        $objWriter->writeCData(Text::controlCharacterPHP2OOXML($shape->getText()));
        $objWriter->endElement();
        // p:sp\p:txBody\a:p\a:r\
        $objWriter->endElement();
        // p:sp\p:txBody\a:p\
        $objWriter->endElement();
        // p:sp\p:txBody\
        $objWriter->endElement();
        // p:sp\
        $objWriter->endElement();
    }

    /**
     * Write chart.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeShapeChart(XMLWriter $objWriter, ShapeChart $shape, int $shapeId): void
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
        $objWriter->writeAttributeIf(0 != $shape->getRotation(), 'rot', CommonDrawing::degreesToAngle($shape->getRotation()));
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
     * Write pic.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeShapePic(XMLWriter $objWriter, AbstractGraphic $shape, int $shapeId): void
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

        if ($shape instanceof AbstractDrawingAdapter && $shape->getExtension() == 'svg') {
            $objWriter->startElement('a:extLst');
            $objWriter->startElement('a:ext');
            $objWriter->writeAttribute('uri', '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}');
            $objWriter->startElement('a16:creationId');
            $objWriter->writeAttribute('xmlns:a16', 'http://schemas.microsoft.com/office/drawing/2014/main');
            $objWriter->writeAttribute('id', '{F8CFD691-5332-EB49-9B42-7D7B3DB9185D}');
            $objWriter->endElement();
            $objWriter->endElement();
            $objWriter->endElement();
        }

        $objWriter->endElement();

        // p:cNvPicPr
        $objWriter->startElement('p:cNvPicPr');

        // a:picLocks
        $objWriter->startElement('a:picLocks');
        $objWriter->writeAttribute('noChangeAspect', '1');
        $objWriter->endElement();

        // #p:cNvPicPr
        $objWriter->endElement();

        // p:nvPr
        $objWriter->startElement('p:nvPr');
        // PlaceHolder
        if ($shape->isPlaceholder()) {
            $objWriter->startElement('p:ph');
            $objWriter->writeAttribute('type', $shape->getPlaceholder()->getType());
            $objWriter->endElement();
        }
        // @link : https://github.com/stefslon/exportToPPTX/blob/master/exportToPPTX.m#L2128
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
            $objWriter->writeAttribute('r:embed', 'rId' . ((int) substr($shape->relationId, strlen('rId')) + 1));
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

        if ($shape instanceof AbstractDrawingAdapter && $shape->getExtension() == 'svg') {
            // a:extLst
            $objWriter->startElement('a:extLst');

            // a:extLst > a:ext
            $objWriter->startElement('a:ext');
            $objWriter->writeAttribute('uri', '{28A0092B-C50C-407E-A947-70E740481C1C}');
            // a:extLst > a:ext > a14:useLocalDpi
            $objWriter->startElement('a14:useLocalDpi');
            $objWriter->writeAttribute('xmlns:a14', 'http://schemas.microsoft.com/office/drawing/2010/main');
            $objWriter->writeAttribute('val', '0');
            // a:extLst > a:ext > ##a14:useLocalDpi
            $objWriter->endElement();
            // a:extLst > ##a:ext
            $objWriter->endElement();

            // a:extLst > a:ext
            $objWriter->startElement('a:ext');
            $objWriter->writeAttribute('uri', '{96DAC541-7B7A-43D3-8B79-37D633B846F1}');
            // a:extLst > a:ext > asvg:svgBlip
            $objWriter->startElement('asvg:svgBlip');
            $objWriter->writeAttribute('xmlns:asvg', 'http://schemas.microsoft.com/office/drawing/2016/SVG/main');
            $objWriter->writeAttribute('r:embed', $shape->relationId);
            // a:extLst > a:ext > ##asvg:svgBlip
            $objWriter->endElement();
            // a:extLst > ##a:ext
            $objWriter->endElement();

            // ##a:extLst
            $objWriter->endElement();
        }

        $objWriter->endElement();

        // a:stretch
        $objWriter->startElement('a:stretch');
        $objWriter->writeElement('a:fillRect');
        $objWriter->endElement();

        $objWriter->endElement();

        // p:spPr
        $objWriter->startElement('p:spPr');
        // a:xfrm
        $objWriter->startElement('a:xfrm');
        $objWriter->writeAttributeIf(0 != $shape->getRotation(), 'rot', CommonDrawing::degreesToAngle($shape->getRotation()));

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
        // // a:prstGeom/a:avLst
        $objWriter->writeElement('a:avLst', null);
        // ##a:prstGeom
        $objWriter->endElement();

        $this->writeFill($objWriter, $shape->getFill());
        $this->writeBorder($objWriter, $shape->getBorder(), '');
        $this->writeShadow($objWriter, $shape->getShadow());

        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write group.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeShapeGroup(XMLWriter $objWriter, Group $group, int &$shapeId): void
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

    protected function writeSlideBackground(AbstractSlideAlias $pSlide, XMLWriter $objWriter): void
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
        // @link : http://www.officeopenxml.com/prSlide-background.php
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
     * Write Transition Slide.
     *
     * @see http://officeopenxml.com/prSlide-transitions.php
     */
    protected function writeSlideTransition(XMLWriter $objWriter, ?Slide\Transition $transition): void
    {
        if (!$transition) {
            return;
        }
        $objWriter->startElement('p:transition');
        if (null !== $transition->getSpeed()) {
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
            case Slide\Transition::TRANSITION_CIRCLE:
                $objWriter->writeElement('p:circle');

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

    private function getGUID(): string
    {
        if (function_exists('com_create_guid')) {
            return com_create_guid();
        }
        mt_srand((int) (microtime(true) * 10000));
        $charid = strtoupper(md5(uniqid((string) mt_rand(), true)));
        $hyphen = chr(45); // "-"
        $uuid = chr(123)// "{"
            . substr($charid, 0, 8) . $hyphen
            . substr($charid, 8, 4) . $hyphen
            . substr($charid, 12, 4) . $hyphen
            . substr($charid, 16, 4) . $hyphen
            . substr($charid, 20, 12)
            . chr(125); // "}"

        return $uuid;
    }
}
