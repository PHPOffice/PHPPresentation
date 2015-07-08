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
use PhpOffice\PhpPresentation\Shape\AbstractDrawing;
use PhpOffice\PhpPresentation\Shape\Chart as ShapeChart;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide as SlideElement;
use PhpOffice\PhpPresentation\Slide\Note;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;

/**
 * Slide writer
 */
class Slide extends AbstractPart
{
    /**
     * Write slide to XML format
     *
     * @param  \PhpOffice\PhpPresentation\Slide $pSlide
     * @return string              XML Output
     * @throws \Exception
     */
    public function writeSlide(SlideElement $pSlide = null)
    {
        // Check slide
        if (is_null($pSlide)) {
            throw new \Exception("Invalid \PhpOffice\PhpPresentation\Slide object passed.");
        }

        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:sld
        $objWriter->startElement('p:sld');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // p:cSld
        $objWriter->startElement('p:cSld');

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
        $shapeId = 0;
        $shapes  = $pSlide->getShapeCollection();
        foreach ($shapes as $shape) {
            // Increment $shapeId
            ++$shapeId;

            // Check type
            if ($shape instanceof RichText) {
                $this->writeShapeText($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Table) {
                $this->writeShapeTable($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Line) {
                $this->writeShapeLine($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof ShapeChart) {
                $this->writeShapeChart($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof AbstractDrawing) {
                $this->writeShapePic($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Group) {
                $this->writeShapeGroup($objWriter, $shape, $shapeId);
            }
        }

        // TODO
        $objWriter->endElement();

        $objWriter->endElement();

        // p:clrMapOvr
        $objWriter->startElement('p:clrMapOvr');

        // a:masterClrMapping
        $objWriter->writeElement('a:masterClrMapping', null);

        $objWriter->endElement();

        if (!is_null($pSlide->getTransition())) {
            $this->writeTransition($objWriter, $pSlide->getTransition());
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write group
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param \PhpOffice\PhpPresentation\Shape\Group $group
     * @param  int $shapeId
     */
    private function writeShapeGroup(XMLWriter $objWriter, Group $group, &$shapeId)
    {
        // p:grpSp
        $objWriter->startElement('p:grpSp');

        // p:nvGrpSpPr
        $objWriter->startElement('p:nvGrpSpPr');

        // p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('name', 'Group '.$shapeId++);
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

        $shapes  = $group->getShapeCollection();
        foreach ($shapes as $shape) {
            // Increment $shapeId
            ++$shapeId;

            // Check type
            if ($shape instanceof RichText) {
                $this->writeShapeText($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Table) {
                $this->writeShapeTable($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Line) {
                $this->writeShapeLine($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof ShapeChart) {
                $this->writeShapeChart($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof AbstractDrawing) {
                $this->writeShapePic($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Group) {
                $this->writeShapeGroup($objWriter, $shape, $shapeId);
            }
        }

        $objWriter->endElement(); // p:grpSp
    }

    /**
     * Write chart
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param \PhpOffice\PhpPresentation\Shape\Chart $shape
     * @param  int $shapeId
     */
    private function writeShapeChart(XMLWriter $objWriter, ShapeChart $shape, $shapeId)
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
        $objWriter->writeElement('p:nvPr', null);

        $objWriter->endElement();

        // p:xfrm
        $objWriter->startElement('p:xfrm');
        $objWriter->writeAttribute('rot', CommonDrawing::degreesToAngle($shape->getRotation()));

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
     * @param  \PhpOffice\Common\XMLWriter  $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\AbstractDrawing $shape
     * @param  int $shapeId
     * @throws \Exception
     */
    private function writeShapePic(XMLWriter $objWriter, AbstractDrawing $shape, $shapeId)
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
        $objWriter->writeElement('p:nvPr', null);
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
        $objWriter->writeAttribute('rot', CommonDrawing::degreesToAngle($shape->getRotation()));

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
     * Write txt
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\RichText   $shape
     * @param  int                            $shapeId
     * @throws \Exception
     */
    private function writeShapeText(XMLWriter $objWriter, RichText $shape, $shapeId)
    {
        // p:sp
        $objWriter->startElement('p:sp');

        // p:sp\p:nvSpPr
        $objWriter->startElement('p:nvSpPr');

        // p:sp\p:nvSpPr\p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', $shapeId);
        $objWriter->writeAttribute('name', '');

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
        $objWriter->writeElement('p:nvPr', null);
        // > p:sp\p:cNvSpPr
        $objWriter->endElement();

        // p:sp\p:spPr
        $objWriter->startElement('p:spPr');

        // p:sp\p:spPr\a:xfrm
        $objWriter->startElement('a:xfrm');
        $objWriter->writeAttribute('rot', CommonDrawing::degreesToAngle($shape->getRotation()));
        
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
        $objWriter->endElement();
        
        if ($shape->getFill()) {
            $this->writeFill($objWriter, $shape->getFill());
        }
        if ($shape->getBorder()->getLineStyle() != Border::LINE_NONE) {
            $this->writeBorder($objWriter, $shape->getBorder(), '');
        }
        if ($shape->getShadow()->isVisible()) {
            $this->writeShadow($objWriter, $shape->getShadow());
        }
        // > p:sp\p:spPr
        $objWriter->endElement();

        // p:txBody
        $objWriter->startElement('p:txBody');

        // a:bodyPr
        //@link :http://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.bodyproperties%28v=office.14%29.aspx
        $objWriter->startElement('a:bodyPr');
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

        $objWriter->endElement();

        // a:lstStyle
        $objWriter->writeElement('a:lstStyle', null);
        
        // Write paragraphs
        $this->writeParagraphs($objWriter, $shape->getParagraphs());

        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write table
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Table      $shape
     * @param  int                            $shapeId
     * @throws \Exception
     */
    private function writeShapeTable(XMLWriter $objWriter, Table $shape, $shapeId)
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
        $objWriter->startElement('p:cNvGraphicFramePr');

        // a:graphicFrameLocks
        $objWriter->startElement('a:graphicFrameLocks');
        $objWriter->writeAttribute('noGrp', '1');
        $objWriter->endElement();

        $objWriter->endElement();

        // p:nvPr
        $objWriter->writeElement('p:nvPr', null);

        $objWriter->endElement();

        // p:xfrm
        $objWriter->startElement('p:xfrm');

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
        $objWriter->writeAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/table');

        // a:tbl
        $objWriter->startElement('a:tbl');

        // a:tblPr
        $objWriter->startElement('a:tblPr');
        $objWriter->writeAttribute('firstRow', '1');
        $objWriter->writeAttribute('bandRow', '1');

        $objWriter->endElement();

        // a:tblGrid
        $objWriter->startElement('a:tblGrid');

        // Write cell widths
        $countCells = count($shape->getRow(0)->getCells());
        for ($cell = 0; $cell < $countCells; $cell++) {
            // a:gridCol
            $objWriter->startElement('a:gridCol');

            // Calculate column width
            $width = $shape->getRow(0)->getCell($cell)->getWidth();
            if ($width == 0) {
                $colCount   = count($shape->getRow(0)->getCells());
                $totalWidth = $shape->getWidth();
                $width      = $totalWidth / $colCount;
            }

            $objWriter->writeAttribute('w', CommonDrawing::pixelsToEmu($width));
            $objWriter->endElement();
        }

        $objWriter->endElement();

        // Colspan / rowspan containers
        $colSpan = array();
        $rowSpan = array();

        // Default border style
        $defaultBorder = new Border();

        // Write rows
        $countRows = count($shape->getRows());
        for ($row = 0; $row < $countRows; $row++) {
            // a:tr
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
                $nextRowBelow  = $shape->getRow($row + 1, true);
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

                // a:bodyPr
                $objWriter->startElement('a:bodyPr');
                $objWriter->writeAttribute('wrap', 'square');
                $objWriter->writeAttribute('rtlCol', '0');

                // a:spAutoFit
                $objWriter->writeElement('a:spAutoFit', null);

                $objWriter->endElement();

                // a:lstStyle
                $objWriter->writeElement('a:lstStyle', null);

                // Write paragraphs
                $this->writeParagraphs($objWriter, $currentCell->getParagraphs());

                $objWriter->endElement();

                // a:tcPr
                $objWriter->startElement('a:tcPr');
                // Alignment (horizontal)
                $firstParagraph  = $currentCell->getParagraph(0);
                $verticalAlign = $firstParagraph->getAlignment()->getVertical();
                if ($verticalAlign != Alignment::VERTICAL_BASE && $verticalAlign != Alignment::VERTICAL_AUTO) {
                    $objWriter->writeAttribute('anchor', $verticalAlign);
                }

                // Determine borders
                $borderLeft         = $currentCell->getBorders()->getLeft();
                $borderRight        = $currentCell->getBorders()->getRight();
                $borderTop          = $currentCell->getBorders()->getTop();
                $borderBottom       = $currentCell->getBorders()->getBottom();
                $borderDiagonalDown = $currentCell->getBorders()->getDiagonalDown();
                $borderDiagonalUp   = $currentCell->getBorders()->getDiagonalUp();

                // Fix PowerPoint implementation
                if (!is_null($nextCellRight) && $nextCellRight->getBorders()->getRight()->getHashCode() != $defaultBorder->getHashCode()) {
                    $borderRight = $nextCellRight->getBorders()->getLeft();
                }
                if (!is_null($nextCellBelow) && $nextCellBelow->getBorders()->getBottom()->getHashCode() != $defaultBorder->getHashCode()) {
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
     * @param  \PhpOffice\Common\XMLWriter           $objWriter  XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\RichText\Paragraph[] $paragraphs
     * @throws \Exception
     */
    private function writeParagraphs(XMLWriter $objWriter, $paragraphs)
    {
        // Loop trough paragraphs
        foreach ($paragraphs as $paragraph) {
            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');
            $objWriter->writeAttribute('algn', $paragraph->getAlignment()->getHorizontal());
            $objWriter->writeAttribute('fontAlgn', $paragraph->getAlignment()->getVertical());
            $objWriter->writeAttribute('marL', CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getMarginLeft()));
            $objWriter->writeAttribute('marR', CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getMarginRight()));
            $objWriter->writeAttribute('indent', CommonDrawing::pixelsToEmu($paragraph->getAlignment()->getIndent()));
            $objWriter->writeAttribute('lvl', $paragraph->getAlignment()->getLevel());

            // Bullet type specified?
            if ($paragraph->getBulletStyle()->getBulletType() != Bullet::TYPE_NONE) {
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
                        // a:rPr
                        $objWriter->startElement('a:rPr');

                        // Bold
                        $objWriter->writeAttribute('b', ($element->getFont()->isBold() ? 'true' : 'false'));

                        // Italic
                        $objWriter->writeAttribute('i', ($element->getFont()->isItalic() ? 'true' : 'false'));

                        // Strikethrough
                        $objWriter->writeAttribute('strike', ($element->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));

                        // Size
                        $objWriter->writeAttribute('sz', ($element->getFont()->getSize() * 100));

                        // Underline
                        $objWriter->writeAttribute('u', $element->getFont()->getUnderline());

                        // Superscript / subscript
                        if ($element->getFont()->isSuperScript() || $element->getFont()->isSubScript()) {
                            if ($element->getFont()->isSuperScript()) {
                                $objWriter->writeAttribute('baseline', '30000');
                            } elseif ($element->getFont()->isSubScript()) {
                                $objWriter->writeAttribute('baseline', '-25000');
                            }
                        }

                        // Color - a:solidFill
                        $objWriter->startElement('a:solidFill');

                        // a:srgbClr
                        $objWriter->startElement('a:srgbClr');
                        $objWriter->writeAttribute('val', $element->getFont()->getColor()->getRGB());
                        $objWriter->endElement();

                        $objWriter->endElement();

                        // Font - a:latin
                        $objWriter->startElement('a:latin');
                        $objWriter->writeAttribute('typeface', $element->getFont()->getName());
                        $objWriter->endElement();

                        // a:hlinkClick
                        if ($element->hasHyperlink()) {
                            $this->writeHyperlink($objWriter, $element);
                        }

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
    private function writeShapeLine(XMLWriter $objWriter, Line $shape, $shapeId)
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
        $objWriter->writeElement('p:nvPr', null);

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
        $objWriter->endElement();

        if ($shape->getBorder()->getLineStyle() != Border::LINE_NONE) {
            $this->writeBorder($objWriter, $shape->getBorder(), '');
        }

        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Border
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter    XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Border     $pBorder      Border
     * @param  string                         $pElementName Element name
     * @throws \Exception
     */
    protected function writeBorder(XMLWriter $objWriter, Border $pBorder, $pElementName = 'L')
    {
        // Line style
        $lineStyle = $pBorder->getLineStyle();
        if ($lineStyle == Border::LINE_NONE) {
            $lineStyle = Border::LINE_SINGLE;
        }

        // Line width
        $lineWidth = 12700 * $pBorder->getLineWidth();

        // a:ln $pElementName
        $objWriter->startElement('a:ln' . $pElementName);
        $objWriter->writeAttribute('w', $lineWidth);
        $objWriter->writeAttribute('cap', 'flat');
        $objWriter->writeAttribute('cmpd', $lineStyle);
        $objWriter->writeAttribute('algn', 'ctr');

        // Fill?
        if ($pBorder->getLineStyle() == Border::LINE_NONE) {
            // a:noFill
            $objWriter->writeElement('a:noFill', null);
        } else {
            // a:solidFill
            $objWriter->startElement('a:solidFill');

            // a:srgbClr
            $objWriter->startElement('a:srgbClr');
            $objWriter->writeAttribute('val', $pBorder->getColor()->getRGB());
            $objWriter->endElement();

            $objWriter->endElement();
        }

        // Dash
        $objWriter->startElement('a:prstDash');
        $objWriter->writeAttribute('val', $pBorder->getDashStyle());
        $objWriter->endElement();

        // a:round
        $objWriter->writeElement('a:round', null);

        // a:headEnd
        $objWriter->startElement('a:headEnd');
        $objWriter->writeAttribute('type', 'none');
        $objWriter->writeAttribute('w', 'med');
        $objWriter->writeAttribute('len', 'med');
        $objWriter->endElement();

        // a:tailEnd
        $objWriter->startElement('a:tailEnd');
        $objWriter->writeAttribute('type', 'none');
        $objWriter->writeAttribute('w', 'med');
        $objWriter->writeAttribute('len', 'med');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Fill
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Fill       $pFill     Fill style
     * @throws \Exception
     */
    protected function writeFill(XMLWriter $objWriter, Fill $pFill)
    {
        // Is it a fill?
        if ($pFill->getFillType() == Fill::FILL_NONE) {
            return;
        }

        // Is it a solid fill?
        if ($pFill->getFillType() == Fill::FILL_SOLID) {
            $this->writeSolidFill($objWriter, $pFill);
            return;
        }

        // Check if this is a pattern type or gradient type
        if ($pFill->getFillType() == Fill::FILL_GRADIENT_LINEAR || $pFill->getFillType() == Fill::FILL_GRADIENT_PATH) {
            // Gradient fill
            $this->writeGradientFill($objWriter, $pFill);
        } else {
            // Pattern fill
            $this->writePatternFill($objWriter, $pFill);
        }
    }

    /**
     * Write Solid Fill
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Fill       $pFill     Fill style
     * @throws \Exception
     */
    protected function writeSolidFill(XMLWriter $objWriter, Fill $pFill)
    {
        // a:gradFill
        $objWriter->startElement('a:solidFill');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getStartColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Gradient Fill
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Fill       $pFill     Fill style
     * @throws \Exception
     */
    protected function writeGradientFill(XMLWriter $objWriter, Fill $pFill)
    {
        // a:gradFill
        $objWriter->startElement('a:gradFill');

        // a:gsLst
        $objWriter->startElement('a:gsLst');
        // a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '0');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getStartColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();

        // a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100000');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getEndColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();

        // a:lin
        $objWriter->startElement('a:lin');
        $objWriter->writeAttribute('ang', CommonDrawing::degreesToAngle($pFill->getRotation()));
        $objWriter->writeAttribute('scaled', '0');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Pattern Fill
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Fill       $pFill     Fill style
     * @throws \Exception
     */
    protected function writePatternFill(XMLWriter $objWriter, Fill $pFill)
    {
        // a:pattFill
        $objWriter->startElement('a:pattFill');

        // fgClr
        $objWriter->startElement('a:fgClr');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getStartColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();

        // bgClr
        $objWriter->startElement('a:bgClr');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getEndColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();
    }

    protected function writeShadow(XMLWriter $objWriter, Shadow $oShadow)
    {
        // a:effectLst
        $objWriter->startElement('a:effectLst');
        
        // a:outerShdw
        $objWriter->startElement('a:outerShdw');
        $objWriter->writeAttribute('blurRad', CommonDrawing::pixelsToEmu($oShadow->getBlurRadius()));
        $objWriter->writeAttribute('dist', CommonDrawing::pixelsToEmu($oShadow->getDistance()));
        $objWriter->writeAttribute('dir', CommonDrawing::degreesToAngle($oShadow->getDirection()));
        $objWriter->writeAttribute('algn', $oShadow->getAlignment());
        $objWriter->writeAttribute('rotWithShape', '0');
        
        // a:srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $oShadow->getColor()->getRGB());
        
        // a:alpha
        $objWriter->startElement('a:alpha');
        $objWriter->writeAttribute('val', $oShadow->getAlpha() * 1000);
        $objWriter->endElement();
        
        $objWriter->endElement();
        
        $objWriter->endElement();
        
        $objWriter->endElement();
    }
    
    /**
     * Write hyperlink
     *
     * @param \PhpOffice\Common\XMLWriter                               $objWriter XML Writer
     * @param \PhpOffice\PhpPresentation\AbstractShape|\PhpOffice\PhpPresentation\Shape\RichText\TextElement $shape
     */
    private function writeHyperlink(XMLWriter $objWriter, $shape)
    {
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
     */
    public function writeNote(Note $pNote = null)
    {
        // Check slide
        if (is_null($pNote)) {
            throw new \Exception("Invalid \PhpOffice\PhpPresentation\Slide\Note object passed.");
        }

        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:notes
        $objWriter->startElement('p:notes');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        // p:cSld
        $objWriter->startElement('p:cSld');

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

        // ## p:nvGrpSpPr
        $objWriter->endElement();

        // p:grpSpPr
        $objWriter->startElement('p:grpSpPr');

        // a:xfrm
        $objWriter->startElement('a:xfrm');

        // a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pNote->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pNote->getOffsetY()));
        $objWriter->endElement(); // a:off

        // a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pNote->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pNote->getExtentY()));
        $objWriter->endElement(); // a:ext

        // a:chOff
        $objWriter->startElement('a:chOff');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pNote->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pNote->getOffsetY()));
        $objWriter->endElement(); // a:chOff

        // a:chExt
        $objWriter->startElement('a:chExt');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pNote->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pNote->getExtentY()));
        $objWriter->endElement(); // a:chExt

        // ## a:xfrm
        $objWriter->endElement();

        // ## p:grpSpPr
        $objWriter->endElement();

        // p:sp
        $objWriter->startElement('p:sp');

        // p:nvSpPr
        $objWriter->startElement('p:nvSpPr');

        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '1');
        $objWriter->writeAttribute('name', 'Notes Placeholder');
        $objWriter->endElement();

        // p:cNvSpPr
        $objWriter->startElement('p:cNvSpPr');

        //a:spLocks
        $objWriter->startElement('a:spLocks');
        $objWriter->writeAttribute('noGrp', '1');
        $objWriter->endElement();

        // ## p:cNvSpPr
        $objWriter->endElement();

        // p:nvPr
        $objWriter->startElement('p:nvPr');

        $objWriter->startElement('p:ph');
        $objWriter->writeAttribute('type', 'body');
        $objWriter->writeAttribute('idx', '1');
        $objWriter->endElement();

        // ## p:nvPr
        $objWriter->endElement();

        // ## p:nvSpPr
        $objWriter->endElement();

        $objWriter->writeElement('p:spPr', null);

        // p:txBody
        $objWriter->startElement('p:txBody');

        $objWriter->writeElement('a:bodyPr', null);
        $objWriter->writeElement('a:lstStyle', null);

        // Loop shapes
        $shapes  = $pNote->getShapeCollection();
        foreach ($shapes as $shape) {
            // Check type
            if ($shape instanceof RichText) {
                $paragraphs = $shape->getParagraphs();
                $this->writeParagraphs($objWriter, $paragraphs);
            }
        }

        // ## p:txBody
        $objWriter->endElement();

        // ## p:sp
        $objWriter->endElement();

        // ## p:spTree
        $objWriter->endElement();

        // ## p:cSld
        $objWriter->endElement();

        // ## p:notes
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write Transition Slide
     * @link http://officeopenxml.com/prSlide-transitions.php
     * @param XMLWriter $objWriter
     * @param Transition $transition
     */
    public function writeTransition(XMLWriter $objWriter, Transition $transition)
    {
        $objWriter->startElement('p:transition');
        if (!is_null($transition->getSpeed())) {
            $objWriter->writeAttribute('spd', $transition->getSpeed());
        }
        $objWriter->writeAttribute('advClick', $transition->hasManualTrigger() ? '1' : '0');
        if ($transition->hasTimeTrigger()) {
            $objWriter->writeAttribute('advTm', $transition->getAdvanceTimeTrigger());
        }

        switch ($transition->getTransitionType()) {
            case Transition::TRANSITION_BLINDS_HORIZONTAL:
                $objWriter->startElement('p:blinds');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_BLINDS_VERTICAL:
                $objWriter->startElement('p:blinds');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_CHECKER_HORIZONTAL:
                $objWriter->startElement('p:checker');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_CHECKER_VERTICAL:
                $objWriter->startElement('p:checker');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_CIRCLE_HORIZONTAL:
                $objWriter->startElement('p:circle');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_CIRCLE_VERTICAL:
                $objWriter->startElement('p:circle');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COMB_HORIZONTAL:
                $objWriter->startElement('p:comb');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COMB_VERTICAL:
                $objWriter->startElement('p:comb');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COVER_DOWN:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'd');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COVER_LEFT:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'l');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COVER_LEFT_DOWN:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'ld');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COVER_LEFT_UP:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'lu');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COVER_RIGHT:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'r');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COVER_RIGHT_DOWN:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'rd');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COVER_RIGHT_UP:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'ru');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_COVER_UP:
                $objWriter->startElement('p:cover');
                $objWriter->writeAttribute('dir', 'u');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_CUT:
                $objWriter->writeElement('p:cut');
                break;
            case Transition::TRANSITION_DIAMOND:
                $objWriter->writeElement('p:diamond');
                break;
            case Transition::TRANSITION_DISSOLVE:
                $objWriter->writeElement('p:dissolve');
                break;
            case Transition::TRANSITION_FADE:
                $objWriter->writeElement('p:fade');
                break;
            case Transition::TRANSITION_NEWSFLASH:
                $objWriter->writeElement('p:newsflash');
                break;
            case Transition::TRANSITION_PLUS:
                $objWriter->writeElement('p:plus');
                break;
            case Transition::TRANSITION_PULL_DOWN:
                $objWriter->startElement('p:pull');
                $objWriter->writeAttribute('dir', 'd');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_PULL_LEFT:
                $objWriter->startElement('p:pull');
                $objWriter->writeAttribute('dir', 'l');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_PULL_RIGHT:
                $objWriter->startElement('p:pull');
                $objWriter->writeAttribute('dir', 'r');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_PULL_UP:
                $objWriter->startElement('p:pull');
                $objWriter->writeAttribute('dir', 'u');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_PUSH_DOWN:
                $objWriter->startElement('p:push');
                $objWriter->writeAttribute('dir', 'd');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_PUSH_LEFT:
                $objWriter->startElement('p:push');
                $objWriter->writeAttribute('dir', 'l');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_PUSH_RIGHT:
                $objWriter->startElement('p:push');
                $objWriter->writeAttribute('dir', 'r');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_PUSH_UP:
                $objWriter->startElement('p:push');
                $objWriter->writeAttribute('dir', 'u');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_RANDOM:
                $objWriter->writeElement('p:random');
                break;
            case Transition::TRANSITION_RANDOMBAR_HORIZONTAL:
                $objWriter->startElement('p:randomBar');
                $objWriter->writeAttribute('dir', 'horz');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_RANDOMBAR_VERTICAL:
                $objWriter->startElement('p:randomBar');
                $objWriter->writeAttribute('dir', 'vert');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_SPLIT_IN_HORIZONTAL:
                $objWriter->startElement('p:split');
                $objWriter->writeAttribute('dir', 'in');
                $objWriter->writeAttribute('orient', 'horz');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_SPLIT_OUT_HORIZONTAL:
                $objWriter->startElement('p:split');
                $objWriter->writeAttribute('dir', 'out');
                $objWriter->writeAttribute('orient', 'horz');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_SPLIT_IN_VERTICAL:
                $objWriter->startElement('p:split');
                $objWriter->writeAttribute('dir', 'in');
                $objWriter->writeAttribute('orient', 'vert');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_SPLIT_OUT_VERTICAL:
                $objWriter->startElement('p:split');
                $objWriter->writeAttribute('dir', 'out');
                $objWriter->writeAttribute('orient', 'vert');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_STRIPS_LEFT_DOWN:
                $objWriter->startElement('p:strips');
                $objWriter->writeAttribute('dir', 'ld');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_STRIPS_LEFT_UP:
                $objWriter->startElement('p:strips');
                $objWriter->writeAttribute('dir', 'lu');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_STRIPS_RIGHT_DOWN:
                $objWriter->startElement('p:strips');
                $objWriter->writeAttribute('dir', 'rd');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_STRIPS_RIGHT_UP:
                $objWriter->startElement('p:strips');
                $objWriter->writeAttribute('dir', 'ru');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_WEDGE:
                $objWriter->writeElement('p:wedge');
                break;
            case Transition::TRANSITION_WIPE_DOWN:
                $objWriter->startElement('p:wipe');
                $objWriter->writeAttribute('dir', 'd');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_WIPE_LEFT:
                $objWriter->startElement('p:wipe');
                $objWriter->writeAttribute('dir', 'l');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_WIPE_RIGHT:
                $objWriter->startElement('p:wipe');
                $objWriter->writeAttribute('dir', 'r');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_WIPE_UP:
                $objWriter->startElement('p:wipe');
                $objWriter->writeAttribute('dir', 'u');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_ZOOM_IN:
                $objWriter->startElement('p:zoom');
                $objWriter->writeAttribute('dir', 'in');
                $objWriter->endElement();
                break;
            case Transition::TRANSITION_ZOOM_OUT:
                $objWriter->startElement('p:zoom');
                $objWriter->writeAttribute('dir', 'out');
                $objWriter->endElement();
                break;
        }

        $objWriter->endElement();
    }
}
