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

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Exception\FileRemoveException;
use PhpOffice\PhpPresentation\Exception\UndefinedChartTypeException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Chart\Legend;
use PhpOffice\PhpPresentation\Shape\Chart\PlotArea;
use PhpOffice\PhpPresentation\Shape\Chart\Title;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Doughnut;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Radar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class PptCharts extends AbstractDecoratorWriter
{
    public function render(): ZipInterface
    {
        for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
            $shape = $this->getDrawingHashTable()->getByIndex($i);
            if ($shape instanceof Chart) {
                $this->getZip()->addFromString('ppt/charts/' . $shape->getIndexedFilename(), $this->writeChart($shape));

                if ($shape->hasIncludedSpreadsheet()) {
                    $this->getZip()->addFromString('ppt/charts/_rels/' . $shape->getIndexedFilename() . '.rels', $this->writeChartRelationships($shape));
                    $pFilename = tempnam(sys_get_temp_dir(), 'PhpSpreadsheet');
                    $this->getZip()->addFromString('ppt/embeddings/' . $shape->getIndexedFilename() . '.xlsx', $this->writeSpreadsheet($this->getPresentation(), $shape, $pFilename . '.xlsx'));

                    // remove temp file
                    if (false === @unlink($pFilename)) {
                        throw new FileRemoveException($pFilename);
                    }
                }
            }
        }

        return $this->getZip();
    }

    /**
     * Write chart to XML format.
     *
     * @return string XML Output
     */
    protected function writeChart(Chart $chart): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // c:chartSpace
        $objWriter->startElement('c:chartSpace');
        $objWriter->writeAttribute('xmlns:c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        // c:date1904
        $objWriter->startElement('c:date1904');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        // c:lang
        $objWriter->startElement('c:lang');
        $objWriter->writeAttribute('val', 'en-US');
        $objWriter->endElement();

        // c:chart
        $objWriter->startElement('c:chart');

        // Title?
        if ($chart->getTitle()->isVisible()) {
            // Write title
            $this->writeTitle($objWriter, $chart->getTitle());
        }

        // c:autoTitleDeleted
        $objWriter->startElement('c:autoTitleDeleted');
        $objWriter->writeAttribute('val', $chart->getTitle()->isVisible() ? '0' : '1');
        $objWriter->endElement();

        // c:view3D
        $objWriter->startElement('c:view3D');

        // c:rotX
        $objWriter->startElement('c:rotX');
        $objWriter->writeAttribute('val', $chart->getView3D()->getRotationX());
        $objWriter->endElement();

        // c:hPercent
        $hPercent = $chart->getView3D()->getHeightPercent();
        $objWriter->writeElementIf(null != $hPercent, 'c:hPercent', 'val', $hPercent);

        // c:rotY
        $objWriter->startElement('c:rotY');
        $objWriter->writeAttribute('val', $chart->getView3D()->getRotationY());
        $objWriter->endElement();

        // c:depthPercent
        $objWriter->startElement('c:depthPercent');
        $objWriter->writeAttribute('val', $chart->getView3D()->getDepthPercent());
        $objWriter->endElement();

        // c:rAngAx
        $objWriter->startElement('c:rAngAx');
        $objWriter->writeAttribute('val', $chart->getView3D()->hasRightAngleAxes() ? '1' : '0');
        $objWriter->endElement();

        // c:perspective
        $objWriter->startElement('c:perspective');
        $objWriter->writeAttribute('val', $chart->getView3D()->getPerspective());
        $objWriter->endElement();

        $objWriter->endElement();

        // Write plot area
        $this->writePlotArea($objWriter, $chart->getPlotArea(), $chart);

        // Legend?
        if ($chart->getLegend()->isVisible()) {
            // Write legend
            $this->writeLegend($objWriter, $chart->getLegend());
        }

        // c:plotVisOnly
        $objWriter->startElement('c:plotVisOnly');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        // c:dispBlanksAs
        $objWriter->startElement('c:dispBlanksAs');
        $objWriter->writeAttribute('val', $chart->getDisplayBlankAs());
        $objWriter->endElement();

        $objWriter->endElement();

        // c:spPr
        $objWriter->startElement('c:spPr');

        // Fill
        $this->writeFill($objWriter, $chart->getFill());

        // Border
        if (Border::LINE_NONE != $chart->getBorder()->getLineStyle()) {
            $this->writeBorder($objWriter, $chart->getBorder(), '');
        }

        // Shadow
        if ($chart->getShadow()->isVisible()) {
            // a:effectLst
            $objWriter->startElement('a:effectLst');

            // a:outerShdw
            $objWriter->startElement('a:outerShdw');
            $objWriter->writeAttribute('blurRad', CommonDrawing::pixelsToEmu($chart->getShadow()->getBlurRadius()));
            $objWriter->writeAttribute('dist', CommonDrawing::pixelsToEmu($chart->getShadow()->getDistance()));
            $objWriter->writeAttribute('dir', CommonDrawing::degreesToAngle((int) $chart->getShadow()->getDirection()));
            $objWriter->writeAttribute('algn', $chart->getShadow()->getAlignment());
            $objWriter->writeAttribute('rotWithShape', '0');

            $this->writeColor($objWriter, $chart->getShadow()->getColor(), $chart->getShadow()->getAlpha());

            $objWriter->endElement();

            $objWriter->endElement();
        }

        $objWriter->endElement();

        // External data?
        if ($chart->hasIncludedSpreadsheet()) {
            // c:externalData
            $objWriter->startElement('c:externalData');
            $objWriter->writeAttribute('r:id', 'rId1');

            // c:autoUpdate
            $objWriter->startElement('c:autoUpdate');
            $objWriter->writeAttribute('val', '0');
            $objWriter->endElement();

            $objWriter->endElement();
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write chart to XML format.
     *
     * @return string String output
     */
    protected function writeSpreadsheet(PhpPresentation $presentation, Chart $chart, string $tempName): string
    {
        // Create new spreadsheet
        $spreadsheet = new Spreadsheet();

        // Set properties
        $title = $chart->getTitle()->getText();
        if (0 == strlen($title)) {
            $title = 'Chart';
        }
        $spreadsheet->getProperties()
            ->setCreator(
                $presentation->getDocumentProperties()->getCreator()
            )
            ->setLastModifiedBy(
                $presentation->getDocumentProperties()->getLastModifiedBy()
            )
            ->setTitle($title);

        // Add chart data
        $sheet = $spreadsheet->setActiveSheetIndex(0);
        $sheet->setTitle('Sheet1');

        // Write series
        $seriesIndex = 0;
        foreach ($chart->getPlotArea()->getType()->getSeries() as $series) {
            // Title
            $sheet->setCellValue(Coordinate::stringFromColumnIndex(2 + $seriesIndex) . 1, $series->getTitle());

            // X-axis
            $axisXData = array_keys($series->getValues());
            $numAxisXData = count($axisXData);
            for ($i = 0; $i < $numAxisXData; ++$i) {
                $sheet->setCellValue('A' . ($i + 2), $axisXData[$i]);
            }

            // Y-axis
            $axisYData = array_values($series->getValues());
            $numAxisYData = count($axisYData);
            for ($i = 0; $i < $numAxisYData; ++$i) {
                $sheet->setCellValue(Coordinate::stringFromColumnIndex(2 + $seriesIndex) . ($i + 2), $axisYData[$i]);
            }

            ++$seriesIndex;
        }

        // Save to string
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($tempName);

        // Load file in memory
        $returnValue = file_get_contents($tempName);
        if (false === @unlink($tempName)) {
            throw new FileRemoveException($tempName);
        }

        return $returnValue;
    }

    /**
     * Write element with value attribute.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeElementWithValAttribute(XMLWriter $objWriter, string $elementName, string $value): void
    {
        $objWriter->startElement($elementName);
        $objWriter->writeAttribute('val', $value);
        $objWriter->endElement();
    }

    /**
     * Write single value or reference.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeSingleValueOrReference(XMLWriter $objWriter, bool $isReference, string $value, string $reference): void
    {
        if (!$isReference) {
            // Value
            $objWriter->writeElement('c:v', $value);

            return;
        }

        // Reference and cache
        // c:strRef
        $objWriter->startElement('c:strRef');
        // c:strRef/c:f
        $objWriter->writeElement('c:f', $reference);
        // c:strRef/c:strCache
        $objWriter->startElement('c:strCache');
        // c:strRef/c:strCache/c:ptCount
        $objWriter->startElement('c:ptCount');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        // c:strRef/c:strCache/c:pt
        $objWriter->startElement('c:pt');
        $objWriter->writeAttribute('idx', '0');
        // c:strRef/c:strCache/c:pt/c:v
        $objWriter->writeElement('c:v', $value);
        // c:strRef/c:strCache/c:pt
        $objWriter->endElement();
        // c:strRef/c:strCache
        $objWriter->endElement();
        // c:strRef
        $objWriter->endElement();
    }

    /**
     * Write series value or reference.
     *
     * @param XMLWriter $objWriter XML Writer
     * @param array<int, mixed> $values
     */
    protected function writeMultipleValuesOrReference(XMLWriter $objWriter, bool $isReference, array $values, string $reference): void
    {
        // c:strLit / c:numLit
        // c:strRef / c:numRef
        $referenceType = ($isReference ? 'Ref' : 'Lit');

        // Get data type from first non-null value
        $dataType = array_reduce($values, function ($carry, $item) {
            if (!isset($item)) {
                return $carry;
            }

            return is_numeric($item) ? 'num' : 'str';
        }, 'num');

        $objWriter->startElement('c:' . $dataType . $referenceType);

        $numValues = count($values);
        if (!$isReference) {
            // Value

            // c:ptCount
            $objWriter->startElement('c:ptCount');
            $objWriter->writeAttribute('val', count($values));
            $objWriter->endElement();

            // Add points
            for ($i = 0; $i < $numValues; ++$i) {
                // c:pt
                $objWriter->startElement('c:pt');
                $objWriter->writeAttribute('idx', $i);
                $objWriter->writeElement('c:v', (string) ($values[$i]));
                $objWriter->endElement();
            }
        } else {
            // Reference
            $objWriter->writeElement('c:f', $reference);
            $objWriter->startElement('c:' . $dataType . 'Cache');

            // c:ptCount
            $objWriter->startElement('c:ptCount');
            $objWriter->writeAttribute('val', count($values));
            $objWriter->endElement();

            // Add points
            for ($i = 0; $i < $numValues; ++$i) {
                // c:pt
                $objWriter->startElement('c:pt');
                $objWriter->writeAttribute('idx', $i);
                $objWriter->writeElement('c:v', (string) ($values[$i]));
                $objWriter->endElement();
            }

            $objWriter->endElement();
        }

        $objWriter->endElement();
    }

    /**
     * Write Title.
     */
    protected function writeTitle(XMLWriter $objWriter, Title $subject): void
    {
        // c:title
        $objWriter->startElement('c:title');

        // c:tx
        $objWriter->startElement('c:tx');

        // c:rich
        $objWriter->startElement('c:rich');

        // a:bodyPr
        $objWriter->writeElement('a:bodyPr', null);

        // a:lstStyle
        $objWriter->writeElement('a:lstStyle', null);

        // a:p
        $objWriter->startElement('a:p');

        // a:pPr
        $objWriter->startElement('a:pPr');
        $objWriter->writeAttribute('algn', $subject->getAlignment()->getHorizontal());
        $objWriter->writeAttribute('fontAlgn', $subject->getAlignment()->getVertical());
        $objWriter->writeAttribute('marL', CommonDrawing::pixelsToEmu($subject->getAlignment()->getMarginLeft()));
        $objWriter->writeAttribute('marR', CommonDrawing::pixelsToEmu($subject->getAlignment()->getMarginRight()));
        $objWriter->writeAttribute('indent', (int) CommonDrawing::pixelsToEmu($subject->getAlignment()->getIndent()));
        $objWriter->writeAttribute('lvl', $subject->getAlignment()->getLevel());

        // a:defRPr
        $objWriter->writeElement('a:defRPr', null);

        $objWriter->endElement();

        // a:r
        $objWriter->startElement('a:r');

        // a:rPr
        $objWriter->startElement('a:rPr');
        $objWriter->writeAttribute('lang', 'en-US');
        $objWriter->writeAttribute('dirty', '0');
        $objWriter->writeAttribute('b', ($subject->getFont()->isBold() ? 'true' : 'false'));
        $objWriter->writeAttribute('i', ($subject->getFont()->isItalic() ? 'true' : 'false'));
        $objWriter->writeAttribute('strike', $subject->getFont()->getStrikethrough());
        $objWriter->writeAttribute('sz', ($subject->getFont()->getSize() * 100));
        $objWriter->writeAttribute('u', $subject->getFont()->getUnderline());
        $objWriter->writeAttributeIf($subject->getFont()->getBaseline() !== 0, 'baseline', $subject->getFont()->getBaseline());
        $objWriter->writeAttribute('cap', $subject->getFont()->getCapitalization());

        // Font - a:solidFill
        $objWriter->startElement('a:solidFill');

        $this->writeColor($objWriter, $subject->getFont()->getColor());

        $objWriter->endElement();

        // Font - a:latin
        $objWriter->startElement('a:latin');
        $objWriter->writeAttribute('typeface', $subject->getFont()->getName());
        $objWriter->endElement();

        $objWriter->endElement();

        // a:t
        $objWriter->writeElement('a:t', $subject->getText());

        $objWriter->endElement();

        // a:endParaRPr
        $objWriter->startElement('a:endParaRPr');
        $objWriter->writeAttribute('lang', 'en-US');
        $objWriter->writeAttribute('dirty', '0');
        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();

        // Write layout
        $this->writeLayout($objWriter, $subject);

        // c:overlay
        $objWriter->startElement('c:overlay');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Plot Area.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writePlotArea(XMLWriter $objWriter, PlotArea $subject, Chart $chart): void
    {
        // c:plotArea
        $objWriter->startElement('c:plotArea');

        // Write layout
        $this->writeLayout($objWriter, $subject);

        // Write chart
        $chartType = $subject->getType();
        if ($chartType instanceof Area) {
            $this->writeTypeArea($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Bar) {
            $this->writeTypeBar($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Bar3D) {
            $this->writeTypeBar3D($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Doughnut) {
            $this->writeTypeDoughnut($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Pie) {
            $this->writeTypePie($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Pie3D) {
            $this->writeTypePie3D($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Line) {
            $this->writeTypeLine($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Radar) {
            $this->writeTypeRadar($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Scatter) {
            $this->writeTypeScatter($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } else {
            throw new UndefinedChartTypeException();
        }

        // Write X axis?
        if ($chartType->hasAxisX()) {
            $this->writeAxis($objWriter, $subject->getAxisX(), Chart\Axis::AXIS_X, $chartType);
        }

        // Write Y axis?
        if ($chartType->hasAxisY()) {
            $this->writeAxis($objWriter, $subject->getAxisY(), Chart\Axis::AXIS_Y, $chartType);
        }

        $objWriter->endElement();
    }

    protected function writeLegend(XMLWriter $objWriter, Legend $subject): void
    {
        // c:legend
        $objWriter->startElement('c:legend');

        // c:legendPos
        $objWriter->startElement('c:legendPos');
        $objWriter->writeAttribute('val', $subject->getPosition());
        $objWriter->endElement();

        // Write layout
        $this->writeLayout($objWriter, $subject);

        // c:overlay
        $objWriter->startElement('c:overlay');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        // c:spPr
        $objWriter->startElement('c:spPr');

        // Fill
        $this->writeFill($objWriter, $subject->getFill());

        // Border
        if (Border::LINE_NONE != $subject->getBorder()->getLineStyle()) {
            $this->writeBorder($objWriter, $subject->getBorder(), '');
        }

        $objWriter->endElement();

        // c:txPr
        $objWriter->startElement('c:txPr');

        // a:bodyPr
        $objWriter->writeElement('a:bodyPr', null);

        // a:lstStyle
        $objWriter->writeElement('a:lstStyle', null);

        // a:p
        $objWriter->startElement('a:p');

        // a:pPr
        $objWriter->startElement('a:pPr');
        $objWriter->writeAttribute('algn', $subject->getAlignment()->getHorizontal());
        $objWriter->writeAttribute('fontAlgn', $subject->getAlignment()->getVertical());
        $objWriter->writeAttribute('marL', CommonDrawing::pixelsToEmu($subject->getAlignment()->getMarginLeft()));
        $objWriter->writeAttribute('marR', CommonDrawing::pixelsToEmu($subject->getAlignment()->getMarginRight()));
        $objWriter->writeAttribute('indent', (int) CommonDrawing::pixelsToEmu($subject->getAlignment()->getIndent()));
        $objWriter->writeAttribute('lvl', $subject->getAlignment()->getLevel());

        // a:defRPr
        $objWriter->startElement('a:defRPr');

        $objWriter->writeAttribute('b', ($subject->getFont()->isBold() ? 'true' : 'false'));
        $objWriter->writeAttribute('i', ($subject->getFont()->isItalic() ? 'true' : 'false'));
        $objWriter->writeAttribute('strike', $subject->getFont()->getStrikethrough());
        $objWriter->writeAttribute('sz', ($subject->getFont()->getSize() * 100));
        $objWriter->writeAttribute('u', $subject->getFont()->getUnderline());
        $objWriter->writeAttributeIf($subject->getFont()->getBaseline() !== 0, 'baseline', $subject->getFont()->getBaseline());

        // Font - a:solidFill
        $objWriter->startElement('a:solidFill');

        $this->writeColor($objWriter, $subject->getFont()->getColor());

        $objWriter->endElement();

        // Font - a:latin
        $objWriter->startElement('a:latin');
        $objWriter->writeAttribute('typeface', $subject->getFont()->getName());
        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();

        // a:endParaRPr
        $objWriter->startElement('a:endParaRPr');
        $objWriter->writeAttribute('lang', 'en-US');
        $objWriter->writeAttribute('dirty', '0');
        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Layout.
     *
     * @param XMLWriter $objWriter XML Writer
     * @param Legend|PlotArea|Title $subject
     */
    protected function writeLayout(XMLWriter $objWriter, $subject): void
    {
        // c:layout
        $objWriter->startElement('c:layout');

        // c:manualLayout
        $objWriter->startElement('c:manualLayout');
        // c:xMode
        $objWriter->startElement('c:xMode');
        $objWriter->writeAttribute('val', 'edge');
        $objWriter->endElement();

        // c:yMode
        $objWriter->startElement('c:yMode');
        $objWriter->writeAttribute('val', 'edge');
        $objWriter->endElement();

        if (0 != $subject->getOffsetX()) {
            // c:x
            $objWriter->startElement('c:x');
            $objWriter->writeAttribute('val', $subject->getOffsetX());
            $objWriter->endElement();
        }

        if (0 != $subject->getOffsetY()) {
            // c:y
            $objWriter->startElement('c:y');
            $objWriter->writeAttribute('val', $subject->getOffsetY());
            $objWriter->endElement();
        }

        if (0 != $subject->getWidth()) {
            // c:w
            $objWriter->startElement('c:w');
            $objWriter->writeAttribute('val', $subject->getWidth());
            $objWriter->endElement();
        }

        if (0 != $subject->getHeight()) {
            // c:h
            $objWriter->startElement('c:h');
            $objWriter->writeAttribute('val', $subject->getHeight());
            $objWriter->endElement();
        }

        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Write Type Area.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeTypeArea(XMLWriter $objWriter, Area $subject, bool $includeSheet = false): void
    {
        // c:lineChart
        $objWriter->startElement('c:areaChart');

        // c:grouping
        $objWriter->startElement('c:grouping');
        $objWriter->writeAttribute('val', 'standard');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:ser > c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:ser > c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:ser > c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // c:ser > c:dLbls
            // @link : https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.charts.areachartseries.aspx
            $objWriter->startElement('c:dLbls');

            // c:ser > c:dLbls > c:showVal
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');

            // c:ser > c:dLbls > c:showCatName
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');

            // c:ser > c:dLbls > c:showSerName
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');

            // c:ser > c:dLbls > c:showPercent
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');

            // c:ser > ##c:dLbls
            $objWriter->endElement();

            if (Fill::FILL_NONE != $series->getFill()->getFillType()) {
                // c:spPr
                $objWriter->startElement('c:spPr');
                // Write fill
                $this->writeFill($objWriter, $series->getFill());
                // ## c:spPr
                $objWriter->endElement();
            }

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:cat
            $objWriter->startElement('c:cat');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:val
            $objWriter->startElement('c:val');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52743552');
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52749440');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Type Bar.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeTypeBar(XMLWriter $objWriter, Bar $subject, bool $includeSheet = false): void
    {
        // c:barChart
        $objWriter->startElement('c:barChart');

        // c:barDir
        $objWriter->startElement('c:barDir');
        $objWriter->writeAttribute('val', $subject->getBarDirection());
        $objWriter->endElement();

        // c:grouping
        $objWriter->startElement('c:grouping');
        $objWriter->writeAttribute('val', $subject->getBarGrouping());
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // Fills for points?
            $dataPointFills = $series->getDataPointFills();
            foreach ($dataPointFills as $key => $value) {
                // c:dPt
                $objWriter->startElement('c:dPt');

                // c:idx
                $this->writeElementWithValAttribute($objWriter, 'c:idx', (string) $key);

                if (Fill::FILL_NONE != $value->getFillType()) {
                    // c:spPr
                    $objWriter->startElement('c:spPr');
                    // Write fill
                    $this->writeFill($objWriter, $value);
                    // ## c:spPr
                    $objWriter->endElement();
                }

                // ## c:dPt
                $objWriter->endElement();
            }

            // c:dLbls
            $objWriter->startElement('c:dLbls');

            if ($series->hasDlblNumFormat()) {
                //c:numFmt
                $objWriter->startElement('c:numFmt');
                $objWriter->writeAttribute('formatCode', $series->getDlblNumFormat());
                $objWriter->writeAttribute('sourceLinked', '0');
                $objWriter->endElement();
            }

            // c:txPr
            $objWriter->startElement('c:txPr');

            // a:bodyPr
            $objWriter->writeElement('a:bodyPr');

            // a:lstStyle
            $objWriter->writeElement('a:lstStyle');

            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');

            // a:defRPr
            $objWriter->startElement('a:defRPr');
            $objWriter->writeAttribute('b', ($series->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($series->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $series->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->getBaseline() !== 0, 'baseline', $series->getFont()->getBaseline());

            // a:solidFill
            $objWriter->startElement('a:solidFill');
            $this->writeColor($objWriter, $series->getFont()->getColor());
            $objWriter->endElement();

            // a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            // >a:defRPr
            $objWriter->endElement();
            // >a:pPr
            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            // >a:p
            $objWriter->endElement();
            // >a:lstStyle
            $objWriter->endElement();

            // c:dLblPos
            $this->writeElementWithValAttribute($objWriter, 'c:dLblPos', $series->getLabelPosition());

            // c:showVal
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');

            // c:showCatName
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');

            // c:showSerName
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');

            // c:showPercent
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');

            // c:separator
            $objWriter->writeElement('c:separator', $series->hasShowSeparator() ? $series->getSeparator() : '');

            // c:showLeaderLines
            $this->writeElementWithValAttribute($objWriter, 'c:showLeaderLines', $series->hasShowLeaderLines() ? '1' : '0');

            $objWriter->endElement();

            // c:spPr
            if (Fill::FILL_NONE != $series->getFill()->getFillType()) {
                // c:spPr
                $objWriter->startElement('c:spPr');
                // Write fill
                $this->writeFill($objWriter, $series->getFill());
                // ## c:spPr
                $objWriter->endElement();
            }

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:cat
            $objWriter->startElement('c:cat');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:val
            $objWriter->startElement('c:val');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:gapWidth
        $objWriter->startElement('c:gapWidth');
        $objWriter->writeAttribute('val', $subject->getGapWidthPercent());
        $objWriter->endElement();

        // c:overlap
        $objWriter->startElement('c:overlap');
        $objWriter->writeAttribute('val', $subject->getOverlapWidthPercent());
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52743552');
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52749440');
        $objWriter->endElement();

        // c:extLst
        $objWriter->startElement('c:extLst');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Type Bar3D.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeTypeBar3D(XMLWriter $objWriter, Bar3D $subject, bool $includeSheet = false): void
    {
        // c:bar3DChart
        $objWriter->startElement('c:bar3DChart');

        // c:barDir
        $objWriter->startElement('c:barDir');
        $objWriter->writeAttribute('val', $subject->getBarDirection());
        $objWriter->endElement();

        // c:grouping
        $objWriter->startElement('c:grouping');
        $objWriter->writeAttribute('val', $subject->getBarGrouping());
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // Fills for points?
            $dataPointFills = $series->getDataPointFills();
            foreach ($dataPointFills as $key => $value) {
                // c:dPt
                $objWriter->startElement('c:dPt');

                // c:idx
                $this->writeElementWithValAttribute($objWriter, 'c:idx', (string) $key);

                if (Fill::FILL_NONE != $value->getFillType()) {
                    // c:spPr
                    $objWriter->startElement('c:spPr');
                    // Write fill
                    $this->writeFill($objWriter, $value);
                    // ## c:spPr
                    $objWriter->endElement();
                }

                // ## c:dPt
                $objWriter->endElement();
            }

            // c:dLbls
            $objWriter->startElement('c:dLbls');

            // c:txPr
            $objWriter->startElement('c:txPr');

            // a:bodyPr
            $objWriter->writeElement('a:bodyPr', null);

            // a:lstStyle
            $objWriter->writeElement('a:lstStyle', null);

            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');

            // a:defRPr
            $objWriter->startElement('a:defRPr');

            $objWriter->writeAttribute('b', ($series->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($series->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $series->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->getBaseline() !== 0, 'baseline', $series->getFont()->getBaseline());

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();
            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // c:showVal
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');

            // c:showCatName
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');

            // c:showSerName
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');

            // c:showPercent
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');

            // c:showLeaderLines
            $this->writeElementWithValAttribute($objWriter, 'c:showLeaderLines', $series->hasShowLeaderLines() ? '1' : '0');

            $objWriter->endElement();

            // c:spPr
            if (Fill::FILL_NONE != $series->getFill()->getFillType()) {
                // c:spPr
                $objWriter->startElement('c:spPr');
                // Write fill
                $this->writeFill($objWriter, $series->getFill());
                // ## c:spPr
                $objWriter->endElement();
            }

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:cat
            $objWriter->startElement('c:cat');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:val
            $objWriter->startElement('c:val');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:gapWidth
        $objWriter->startElement('c:gapWidth');
        $objWriter->writeAttribute('val', $subject->getGapWidthPercent());
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52743552');
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52749440');
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Type Pie.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeTypeDoughnut(XMLWriter $objWriter, Doughnut $subject, bool $includeSheet = false): void
    {
        // c:pieChart
        $objWriter->startElement('c:doughnutChart');

        // c:varyColors
        $objWriter->startElement('c:varyColors');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // Fills for points?
            $dataPointFills = $series->getDataPointFills();
            foreach ($dataPointFills as $key => $value) {
                // c:dPt
                $objWriter->startElement('c:dPt');
                $this->writeElementWithValAttribute($objWriter, 'c:idx', (string) $key);
                // c:dPt/c:spPr
                $objWriter->startElement('c:spPr');
                $this->writeFill($objWriter, $value);
                // c:dPt/##c:spPr
                $objWriter->endElement();
                // ##c:dPt
                $objWriter->endElement();
            }

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:cat
            $objWriter->startElement('c:cat');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:val
            $objWriter->startElement('c:val');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        if (isset($series) && is_object($series) && $series instanceof Chart\Series) {
            // c:dLbls
            $objWriter->startElement('c:dLbls');

            if ($series->hasDlblNumFormat()) {
                //c:numFmt
                $objWriter->startElement('c:numFmt');
                $objWriter->writeAttribute('formatCode', $series->getDlblNumFormat());
                $objWriter->writeAttribute('sourceLinked', '0');
                $objWriter->endElement();
            }

            // c:dLbls\c:txPr
            $objWriter->startElement('c:txPr');
            $objWriter->writeElement('a:bodyPr', null);
            $objWriter->writeElement('a:lstStyle', null);

            // c:dLbls\c:txPr\a:p
            $objWriter->startElement('a:p');

            // c:dLbls\c:txPr\a:p\a:pPr
            $objWriter->startElement('a:pPr');

            // c:dLbls\c:txPr\a:p\a:pPr\a:defRPr
            $objWriter->startElement('a:defRPr');
            $objWriter->writeAttribute('b', ($series->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($series->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $series->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->getBaseline() !== 0, 'baseline', $series->getFont()->getBaseline());

            // c:dLbls\c:txPr\a:p\a:pPr\a:defRPr\a:solidFill
            $objWriter->startElement('a:solidFill');
            $this->writeColor($objWriter, $series->getFont()->getColor());
            $objWriter->endElement();

            // c:dLbls\c:txPr\a:p\a:pPr\a:defRPr\a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();
            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            // c:dLbls\c:txPr\a:p\a:pPr\a:defRPr\
            $objWriter->endElement();
            // c:dLbls\c:txPr\a:p\a:pPr\
            $objWriter->endElement();

            // c:dLbls\c:txPr\a:p\a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            // c:dLbls\c:txPr\a:p\
            $objWriter->endElement();
            // c:dLbls\c:txPr\
            $objWriter->endElement();

            $this->writeElementWithValAttribute($objWriter, 'c:showLegendKey', $series->hasShowLegendKey() ? '1' : '0');
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');
            $this->writeElementWithValAttribute($objWriter, 'c:showBubbleSize', '0');
            $this->writeElementWithValAttribute($objWriter, 'c:showLeaderLines', $series->hasShowLeaderLines() ? '1' : '0');

            $separator = $series->getSeparator();
            if (!empty($separator) && PHP_EOL != $separator) {
                // c:dLbls\c:separator
                $objWriter->writeElement('c:separator', $separator);
            }

            // c:dLbls\
            $objWriter->endElement();
        }

        $this->writeElementWithValAttribute($objWriter, 'c:firstSliceAng', '0');
        $this->writeElementWithValAttribute($objWriter, 'c:holeSize', (string) $subject->getHoleSize());

        $objWriter->endElement();
    }

    /**
     * Write Type Pie.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeTypePie(XMLWriter $objWriter, Pie $subject, bool $includeSheet = false): void
    {
        // c:pieChart
        $objWriter->startElement('c:pieChart');

        // c:varyColors
        $objWriter->startElement('c:varyColors');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // Fills for points?
            $dataPointFills = $series->getDataPointFills();
            foreach ($dataPointFills as $key => $value) {
                // c:dPt
                $objWriter->startElement('c:dPt');
                $this->writeElementWithValAttribute($objWriter, 'c:idx', (string) $key);
                // c:dPt/c:spPr
                $objWriter->startElement('c:spPr');
                $this->writeFill($objWriter, $value);
                // c:dPt/##c:spPr
                $objWriter->endElement();
                // ##c:dPt
                $objWriter->endElement();
            }

            // c:dLbls
            $objWriter->startElement('c:dLbls');

            if ($series->hasDlblNumFormat()) {
                //c:numFmt
                $objWriter->startElement('c:numFmt');
                $objWriter->writeAttribute('formatCode', $series->getDlblNumFormat());
                $objWriter->writeAttribute('sourceLinked', '0');
                $objWriter->endElement();
            }

            // c:txPr
            $objWriter->startElement('c:txPr');

            // a:bodyPr
            $objWriter->writeElement('a:bodyPr', null);

            // a:lstStyle
            $objWriter->writeElement('a:lstStyle', null);

            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');

            // a:defRPr
            $objWriter->startElement('a:defRPr');

            $objWriter->writeAttribute('b', ($series->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($series->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $series->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->getBaseline() !== 0, 'baseline', $series->getFont()->getBaseline());

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();
            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // c:dLblPos
            $this->writeElementWithValAttribute($objWriter, 'c:dLblPos', $series->getLabelPosition());

            // c:showLegendKey
            $this->writeElementWithValAttribute($objWriter, 'c:showLegendKey', $series->hasShowLegendKey() ? '1' : '0');

            // c:showVal
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');

            // c:showCatName
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');

            // c:showSerName
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');

            // c:showPercent
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');

            // c:showLeaderLines
            $this->writeElementWithValAttribute($objWriter, 'c:showLeaderLines', $series->hasShowLeaderLines() ? '1' : '0');

            $objWriter->endElement();

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:cat
            $objWriter->startElement('c:cat');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:val
            $objWriter->startElement('c:val');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        $objWriter->endElement();
    }

    /**
     * Write Type Pie3D.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeTypePie3D(XMLWriter $objWriter, Pie3D $subject, bool $includeSheet = false): void
    {
        // c:pie3DChart
        $objWriter->startElement('c:pie3DChart');

        // c:varyColors
        $objWriter->startElement('c:varyColors');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // c:explosion
            $objWriter->startElement('c:explosion');
            $objWriter->writeAttribute('val', $subject->getExplosion());
            $objWriter->endElement();

            // Fills for points?
            $dataPointFills = $series->getDataPointFills();
            foreach ($dataPointFills as $key => $value) {
                // c:dPt
                $objWriter->startElement('c:dPt');
                $this->writeElementWithValAttribute($objWriter, 'c:idx', (string) $key);
                // c:dPt/c:spPr
                $objWriter->startElement('c:spPr');
                $this->writeFill($objWriter, $value);
                // c:dPt/##c:spPr
                $objWriter->endElement();
                // ##c:dPt
                $objWriter->endElement();
            }

            // c:dLbls
            $objWriter->startElement('c:dLbls');

            // c:txPr
            $objWriter->startElement('c:txPr');

            // a:bodyPr
            $objWriter->writeElement('a:bodyPr', null);

            // a:lstStyle
            $objWriter->writeElement('a:lstStyle', null);

            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');

            // a:defRPr
            $objWriter->startElement('a:defRPr');

            $objWriter->writeAttribute('b', ($series->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($series->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $series->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->getBaseline() !== 0, 'baseline', $series->getFont()->getBaseline());

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();
            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // c:dLblPos
            $this->writeElementWithValAttribute($objWriter, 'c:dLblPos', $series->getLabelPosition());

            // c:showVal
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');

            // c:showCatName
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');

            // c:showSerName
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');

            // c:showPercent
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');

            // c:showLeaderLines
            $this->writeElementWithValAttribute($objWriter, 'c:showLeaderLines', $series->hasShowLeaderLines() ? '1' : '0');

            $objWriter->endElement();

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:cat
            $objWriter->startElement('c:cat');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:val
            $objWriter->startElement('c:val');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        $objWriter->endElement();
    }

    /**
     * Write Type Line.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeTypeLine(XMLWriter $objWriter, Line $subject, bool $includeSheet = false): void
    {
        // c:lineChart
        $objWriter->startElement('c:lineChart');

        // c:grouping
        $objWriter->startElement('c:grouping');
        $objWriter->writeAttribute('val', 'standard');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // c:spPr
            $objWriter->startElement('c:spPr');
            // Write fill
            $this->writeFill($objWriter, $series->getFill());
            // Write outline
            $this->writeOutline($objWriter, $series->getOutline());
            // ## c:spPr
            $objWriter->endElement();

            // Marker
            $this->writeSeriesMarker($objWriter, $series->getMarker());

            // c:dLbls
            $objWriter->startElement('c:dLbls');

            // c:txPr
            $objWriter->startElement('c:txPr');

            // a:bodyPr
            $objWriter->writeElement('a:bodyPr', null);

            // a:lstStyle
            $objWriter->writeElement('a:lstStyle', null);

            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');

            // a:defRPr
            $objWriter->startElement('a:defRPr');

            $objWriter->writeAttribute('b', ($series->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($series->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $series->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->getBaseline() !== 0, 'baseline', $series->getFont()->getBaseline());

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();
            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // c:dLblPos
            $this->writeElementWithValAttribute($objWriter, 'c:dLblPos', $series->getLabelPosition());

            // c:showVal
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');

            // c:showCatName
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');

            // c:showSerName
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');

            // c:showPercent
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');

            // c:showLeaderLines
            $this->writeElementWithValAttribute($objWriter, 'c:showLeaderLines', $series->hasShowLeaderLines() ? '1' : '0');

            // > c:dLbls
            $objWriter->endElement();

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:cat
            $objWriter->startElement('c:cat');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:val
            $objWriter->startElement('c:val');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            // c:smooth
            $objWriter->startElement('c:smooth');
            $objWriter->writeAttribute('val', $subject->isSmooth() ? '1' : '0');
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:marker
        $objWriter->startElement('c:marker');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52743552');
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52749440');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Type Radar.
     *
     * @param XMLWriter $objWriter XML Writer
     */
    protected function writeTypeRadar(XMLWriter $objWriter, Radar $subject, bool $includeSheet = false): void
    {
        // c:scatterChart
        $objWriter->startElement('c:radarChart');

        // c:radarStyle
        $objWriter->startElement('c:radarStyle');
        $objWriter->writeAttribute('val', 'marker');
        $objWriter->endElement();

        // c:varyColors
        $objWriter->startElement('c:varyColors');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // c:spPr
            $objWriter->startElement('c:spPr');
            // Write fill
            $this->writeFill($objWriter, $series->getFill());
            // Write outline
            $this->writeOutline($objWriter, $series->getOutline());
            // ## c:spPr
            $objWriter->endElement();

            // Marker
            $this->writeSeriesMarker($objWriter, $series->getMarker());

            // c:dLbls
            $objWriter->startElement('c:dLbls');

            // c:txPr
            $objWriter->startElement('c:txPr');

            // a:bodyPr
            $objWriter->writeElement('a:bodyPr', null);

            // a:lstStyle
            $objWriter->writeElement('a:lstStyle', null);

            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');

            // a:defRPr
            $objWriter->startElement('a:defRPr');

            $objWriter->writeAttribute('b', ($series->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($series->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $series->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->getBaseline() !== 0, 'baseline', $series->getFont()->getBaseline());

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();
            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // c:showLegendKey
            $this->writeElementWithValAttribute($objWriter, 'c:showLegendKey', $series->hasShowLegendKey() ? '1' : '0');

            // c:showVal
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');

            // c:showCatName
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');

            // c:showSerName
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');

            // c:showPercent
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');

            // c:showLeaderLines
            $this->writeElementWithValAttribute($objWriter, 'c:showLeaderLines', $series->hasShowLeaderLines() ? '1' : '0');

            $objWriter->endElement();

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:cat
            $objWriter->startElement('c:cat');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:val
            $objWriter->startElement('c:val');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52743552');
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52749440');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Type Scatter.
     */
    protected function writeTypeScatter(XMLWriter $objWriter, Scatter $subject, bool $includeSheet = false): void
    {
        // c:scatterChart
        $objWriter->startElement('c:scatterChart');

        // c:scatterStyle
        $objWriter->startElement('c:scatterStyle');
        $objWriter->writeAttribute('val', 'lineMarker');
        $objWriter->endElement();

        // c:varyColors
        $objWriter->startElement('c:varyColors');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getSeries() as $series) {
            // c:ser
            $objWriter->startElement('c:ser');

            // c:idx
            $objWriter->startElement('c:idx');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:order
            $objWriter->startElement('c:order');
            $objWriter->writeAttribute('val', $seriesIndex);
            $objWriter->endElement();

            // c:tx
            $objWriter->startElement('c:tx');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex(2 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // c:spPr
            $objWriter->startElement('c:spPr');
            // Write fill
            $this->writeFill($objWriter, $series->getFill());
            // Write outline
            $this->writeOutline($objWriter, $series->getOutline());
            // ## c:spPr
            $objWriter->endElement();

            // Marker
            $this->writeSeriesMarker($objWriter, $series->getMarker());

            // c:dLbls
            $objWriter->startElement('c:dLbls');

            // c:txPr
            $objWriter->startElement('c:txPr');

            // a:bodyPr
            $objWriter->writeElement('a:bodyPr', null);

            // a:lstStyle
            $objWriter->writeElement('a:lstStyle', null);

            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');

            // a:defRPr
            $objWriter->startElement('a:defRPr');

            $objWriter->writeAttribute('b', ($series->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($series->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $series->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->getBaseline() !== 0, 'baseline', $series->getFont()->getBaseline());

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();
            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $series->getFont()->getName());
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // c:showLegendKey
            $this->writeElementWithValAttribute($objWriter, 'c:showLegendKey', $series->hasShowLegendKey() ? '1' : '0');

            // c:showVal
            $this->writeElementWithValAttribute($objWriter, 'c:showVal', $series->hasShowValue() ? '1' : '0');

            // c:showCatName
            $this->writeElementWithValAttribute($objWriter, 'c:showCatName', $series->hasShowCategoryName() ? '1' : '0');

            // c:showSerName
            $this->writeElementWithValAttribute($objWriter, 'c:showSerName', $series->hasShowSeriesName() ? '1' : '0');

            // c:showPercent
            $this->writeElementWithValAttribute($objWriter, 'c:showPercent', $series->hasShowPercentage() ? '1' : '0');

            // c:separator
            $separator = $series->getSeparator();
            if (!empty($separator) && PHP_EOL != $separator) {
                // c:dLbls\c:separator
                $objWriter->writeElement('c:separator', $separator);
            }

            // c:showLeaderLines
            $this->writeElementWithValAttribute($objWriter, 'c:showLeaderLines', $series->hasShowLeaderLines() ? '1' : '0');

            $objWriter->endElement();

            // Write X axis data
            $axisXData = array_keys($series->getValues());

            // c:xVal
            $objWriter->startElement('c:xVal');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisXData, 'Sheet1!$A$2:$A$' . (1 + count($axisXData)));
            $objWriter->endElement();

            // Write Y axis data
            $axisYData = array_values($series->getValues());

            // c:yVal
            $objWriter->startElement('c:yVal');
            $coords = ($includeSheet ? 'Sheet1!$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$2:$' . Coordinate::stringFromColumnIndex($seriesIndex + 2) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            // c:smooth
            $objWriter->startElement('c:smooth');
            $objWriter->writeAttribute('val', $subject->isSmooth() ? '1' : '0');
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52743552');
        $objWriter->endElement();

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '52749440');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write chart relationships to XML format.
     *
     * @return string XML Output
     */
    protected function writeChartRelationships(Chart $pChart): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

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

    protected function writeSeriesMarker(XMLWriter $objWriter, Chart\Marker $marker): void
    {
        // c:marker
        $objWriter->startElement('c:marker');
        // c:marker > c:symbol
        $objWriter->startElement('c:symbol');
        $objWriter->writeAttribute('val', $marker->getSymbol());
        $objWriter->endElement();

        // Size if different of none
        if (Chart\Marker::SYMBOL_NONE != $marker->getSymbol()) {
            $markerSize = (int) $marker->getSize();
            if ($markerSize < 2) {
                $markerSize = 2;
            }
            if ($markerSize > 72) {
                $markerSize = 72;
            }

            /*
             * c:marker > c:size
             * Size in points
             * @link : https://msdn.microsoft.com/en-us/library/hh658135(v=office.12).aspx
             */
            $objWriter->startElement('c:size');
            $objWriter->writeAttribute('val', $markerSize);
            $objWriter->endElement();
        }

        // // c:marker > c:spPr
        $objWriter->startElement('c:spPr');
        $this->writeFill($objWriter, $marker->getFill());
        $this->writeBorder($objWriter, $marker->getBorder(), '', true);
        $objWriter->endElement();

        // > c:marker
        $objWriter->endElement();
    }

    protected function writeAxis(XMLWriter $objWriter, Chart\Axis $oAxis, string $typeAxis, Chart\Type\AbstractType $typeChart): void
    {
        if (Chart\Axis::AXIS_X != $typeAxis && Chart\Axis::AXIS_Y != $typeAxis) {
            return;
        }

        $crossesAt = $oAxis->getCrossesAt();
        $orientation = $oAxis->isReversedOrder() ? 'maxMin' : 'minMax';

        if (Chart\Axis::AXIS_X == $typeAxis) {
            $mainElement = 'c:catAx';
            $axIdVal = '52743552';
            $axPosVal = $crossesAt === 'max' ? 't' : 'b';
            $crossAxVal = '52749440';
        } else {
            $mainElement = 'c:valAx';
            $axIdVal = '52749440';
            $axPosVal = $crossesAt === 'max' ? 'r' : 'l';
            $crossAxVal = '52743552';
        }

        // $mainElement
        $objWriter->startElement($mainElement);

        // $mainElement > c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', $axIdVal);
        $objWriter->endElement();

        // $mainElement > c:scaling
        $objWriter->startElement('c:scaling');

        // $mainElement > c:scaling > c:orientation
        $objWriter->startElement('c:orientation');
        $objWriter->writeAttribute('val', $orientation);
        $objWriter->endElement();

        if (null !== $oAxis->getMaxBounds()) {
            $objWriter->startElement('c:max');
            $objWriter->writeAttribute('val', $oAxis->getMaxBounds());
            $objWriter->endElement();
        }

        if (null !== $oAxis->getMinBounds()) {
            $objWriter->startElement('c:min');
            $objWriter->writeAttribute('val', $oAxis->getMinBounds());
            $objWriter->endElement();
        }

        // $mainElement > ##c:scaling
        $objWriter->endElement();

        // $mainElement > c:delete
        $objWriter->startElement('c:delete');
        $objWriter->writeAttribute('val', $oAxis->isVisible() ? '0' : '1');
        $objWriter->endElement();

        // $mainElement > c:axPos
        $objWriter->startElement('c:axPos');
        $objWriter->writeAttribute('val', $axPosVal);
        $objWriter->endElement();

        $oMajorGridlines = $oAxis->getMajorGridlines();
        if ($oMajorGridlines instanceof Gridlines) {
            $objWriter->startElement('c:majorGridlines');

            $this->writeAxisGridlines($objWriter, $oMajorGridlines);

            $objWriter->endElement();
        }

        $oMinorGridlines = $oAxis->getMinorGridlines();
        if ($oMinorGridlines instanceof Gridlines) {
            $objWriter->startElement('c:minorGridlines');

            $this->writeAxisGridlines($objWriter, $oMinorGridlines);

            $objWriter->endElement();
        }

        if ('' != $oAxis->getTitle()) {
            // c:title
            $objWriter->startElement('c:title');

            // c:tx
            $objWriter->startElement('c:tx');

            // c:rich
            $objWriter->startElement('c:rich');

            // a:bodyPr
            $objWriter->startElement('a:bodyPr');
            $objWriter->writeAttributeIf($oAxis->getTitleRotation() != 0, 'rot', CommonDrawing::degreesToAngle((int) $oAxis->getTitleRotation()));
            $objWriter->endElement();

            // a:lstStyle
            $objWriter->writeElement('a:lstStyle', null);

            // a:p
            $objWriter->startElement('a:p');

            // a:pPr
            $objWriter->startElement('a:pPr');

            // a:defRPr
            $objWriter->startElement('a:defRPr');

            $objWriter->writeAttribute('b', ($oAxis->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($oAxis->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', $oAxis->getFont()->getStrikethrough());
            $objWriter->writeAttribute('sz', ($oAxis->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $oAxis->getFont()->getUnderline());
            $objWriter->writeAttributeIf($oAxis->getFont()->getBaseline() !== 0, 'baseline', $oAxis->getFont()->getBaseline());

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');
            $this->writeColor($objWriter, $oAxis->getFont()->getColor());
            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
            $objWriter->writeAttribute('typeface', $oAxis->getFont()->getName());
            $objWriter->endElement();
            // a:ea
            $objWriter->startElement('a:ea');
            $objWriter->writeAttribute('typeface', $oAxis->getFont()->getName());
            $objWriter->endElement();

            $objWriter->endElement();

            // ## a:pPr
            $objWriter->endElement();

            // a:r
            $objWriter->startElement('a:r');

            // a:rPr
            $objWriter->startElement('a:rPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            // a:t
            $objWriter->writeElement('a:t', $oAxis->getTitle());

            // ## a:r
            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            // ## a:p
            $objWriter->endElement();

            // ## c:rich
            $objWriter->endElement();

            // ## c:tx
            $objWriter->endElement();

            // ## c:title
            $objWriter->endElement();
        }

        // c:numFmt
        $objWriter->startElement('c:numFmt');
        $objWriter->writeAttribute('formatCode', $oAxis->getFormatCode());
        $objWriter->writeAttribute('sourceLinked', '1');
        $objWriter->endElement();

        // c:majorTickMark
        $objWriter->startElement('c:majorTickMark');
        $objWriter->writeAttribute('val', $oAxis->getMajorTickMark());
        $objWriter->endElement();

        // c:minorTickMark
        $objWriter->startElement('c:minorTickMark');
        $objWriter->writeAttribute('val', $oAxis->getMinorTickMark());
        $objWriter->endElement();

        // c:tickLblPos
        $objWriter->startElement('c:tickLblPos');
        $objWriter->writeAttribute('val', $oAxis->getTickLabelPosition());
        $objWriter->endElement();

        // c:spPr
        $objWriter->startElement('c:spPr');
        $this->writeOutline($objWriter, $oAxis->getOutline());
        $objWriter->endElement();

        // c:txPr
        $objWriter->startElement('c:txPr');

        // a:bodyPr
        $objWriter->writeElement('a:bodyPr', null);

        // a:lstStyle
        $objWriter->writeElement('a:lstStyle', null);

        // a:p
        $objWriter->startElement('a:p');

        // a:pPr
        $objWriter->startElement('a:pPr');

        // a:defRPr
        $objWriter->startElement('a:defRPr');
        $objWriter->writeAttribute('b', ($oAxis->getTickLabelFont()->isBold() ? 'true' : 'false'));
        $objWriter->writeAttribute('i', ($oAxis->getTickLabelFont()->isItalic() ? 'true' : 'false'));
        $objWriter->writeAttribute('strike', $oAxis->getFont()->getStrikethrough());
        $objWriter->writeAttribute('sz', ($oAxis->getTickLabelFont()->getSize() * 100));
        $objWriter->writeAttribute('u', $oAxis->getTickLabelFont()->getUnderline());
        $objWriter->writeAttributeIf($oAxis->getTickLabelFont()->getBaseline() !== 0, 'baseline', $oAxis->getTickLabelFont()->getBaseline());

        // Font - a:solidFill
        $objWriter->startElement('a:solidFill');
        $this->writeColor($objWriter, $oAxis->getTickLabelFont()->getColor());
        $objWriter->endElement();

        // Font - a:latin
        $objWriter->startElement('a:latin');
        $objWriter->writeAttribute('typeface', $oAxis->getTickLabelFont()->getName());
        $objWriter->endElement();

        // Font - a:ea
        $objWriter->startElement('a:ea');
        $objWriter->writeAttribute('typeface', $oAxis->getTickLabelFont()->getName());
        $objWriter->endElement();

        //## a:defRPr
        $objWriter->endElement();

        //## a:pPr
        $objWriter->endElement();

        // a:endParaRPr
        $objWriter->startElement('a:endParaRPr');
        $objWriter->writeAttribute('lang', 'en-US');
        $objWriter->writeAttribute('dirty', '0');
        $objWriter->endElement();

        // ## a:p
        $objWriter->endElement();

        // ## c:txPr
        $objWriter->endElement();

        // c:crossAx
        $objWriter->startElement('c:crossAx');
        $objWriter->writeAttribute('val', $crossAxVal);
        $objWriter->endElement();

        // c:crosses "autoZero" | "min" | "max" | custom string value
        if (in_array($crossesAt, ['autoZero', 'min', 'max'])) {
            $objWriter->startElement('c:crosses');
            $objWriter->writeAttribute('val', $crossesAt);
            $objWriter->endElement();
        } else {
            $objWriter->startElement('c:crossesAt');
            $objWriter->writeAttribute('val', $crossesAt);
            $objWriter->endElement();
        }

        if (Chart\Axis::AXIS_X == $typeAxis) {
            // c:lblAlgn
            $objWriter->startElement('c:lblAlgn');
            $objWriter->writeAttribute('val', 'ctr');
            $objWriter->endElement();

            // c:lblOffset
            $objWriter->startElement('c:lblOffset');
            $objWriter->writeAttribute('val', '100');
            $objWriter->endElement();

            // c:majorUnit
            if ($oAxis->getMajorUnit() !== null) {
                $objWriter->startElement('c:tickLblSkip');
                $objWriter->writeAttribute('val', $oAxis->getMajorUnit());
                $objWriter->endElement();
            }
        }

        if (Chart\Axis::AXIS_Y == $typeAxis) {
            // c:crossBetween
            $objWriter->startElement('c:crossBetween');
            // midCat : Position Axis On Tick Marks
            // between : Between Tick Marks
            if ($typeChart instanceof Area) {
                $objWriter->writeAttribute('val', 'midCat');
            } else {
                $objWriter->writeAttribute('val', 'between');
            }
            $objWriter->endElement();

            // c:majorUnit
            if ($oAxis->getMajorUnit() !== null) {
                $objWriter->startElement('c:majorUnit');
                $objWriter->writeAttribute('val', $oAxis->getMajorUnit());
                $objWriter->endElement();
            }

            // c:minorUnit
            if ($oAxis->getMinorUnit() !== null) {
                $objWriter->startElement('c:minorUnit');
                $objWriter->writeAttribute('val', $oAxis->getMinorUnit());
                $objWriter->endElement();
            }
        }

        $objWriter->endElement();
    }

    protected function writeAxisGridlines(XMLWriter $objWriter, Gridlines $oGridlines): void
    {
        // c:spPr
        $objWriter->startElement('c:spPr');

        // Outline
        $this->writeOutline($objWriter, $oGridlines->getOutline());

        // ##c:spPr
        $objWriter->endElement();
    }
}
