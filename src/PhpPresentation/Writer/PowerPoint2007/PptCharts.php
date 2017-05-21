<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Chart\Legend;
use PhpOffice\PhpPresentation\Shape\Chart\PlotArea;
use PhpOffice\PhpPresentation\Shape\Chart\Title;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;

class PptCharts extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     * @throws \Exception
     */
    public function render()
    {
        for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
            $shape = $this->getDrawingHashTable()->getByIndex($i);
            if ($shape instanceof Chart) {
                $this->getZip()->addFromString('ppt/charts/' . $shape->getIndexedFilename(), $this->writeChart($shape));

                if ($shape->hasIncludedSpreadsheet()) {
                    $this->getZip()->addFromString('ppt/charts/_rels/' . $shape->getIndexedFilename() . '.rels', $this->writeChartRelationships($shape));
                    $pFilename = tempnam(sys_get_temp_dir(), 'PHPExcel');
                    $this->getZip()->addFromString('ppt/embeddings/' . $shape->getIndexedFilename() . '.xlsx', $this->writeSpreadsheet($this->getPresentation(), $shape, $pFilename . '.xlsx'));
                }
            }
        }
        return $this->getZip();
    }


    /**
     * Write chart to XML format
     *
     * @param  \PhpOffice\PhpPresentation\Shape\Chart $chart
     * @return string                    XML Output
     * @throws \Exception
     */
    public function writeChart(Chart $chart)
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
        $objWriter->writeElementIf($hPercent != null, 'c:hPercent', 'val', $hPercent . '%');

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

        $objWriter->endElement();

        // c:spPr
        $objWriter->startElement('c:spPr');

        // Fill
        $this->writeFill($objWriter, $chart->getFill());

        // Border
        if ($chart->getBorder()->getLineStyle() != Border::LINE_NONE) {
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
            $objWriter->writeAttribute('dir', CommonDrawing::degreesToAngle($chart->getShadow()->getDirection()));
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
     * Write chart to XML format
     *
     * @param  PhpPresentation $presentation
     * @param  \PhpOffice\PhpPresentation\Shape\Chart $chart
     * @param  string $tempName
     * @return string                    String output
     * @throws \Exception
     */
    public function writeSpreadsheet(PhpPresentation $presentation, $chart, $tempName)
    {
        // Need output?
        if (!$chart->hasIncludedSpreadsheet()) {
            throw new \Exception('No spreadsheet output is required for the given chart.');
        }

        // Verify PHPExcel
        if (!class_exists('PHPExcel')) {
            throw new \Exception('PHPExcel has not been loaded. Include PHPExcel.php in your script, e.g. require_once \'PHPExcel.php\'.');
        }

        // Create new spreadsheet
        $workbook = new \PHPExcel();

        // Set properties
        $title = $chart->getTitle()->getText();
        if (strlen($title) == 0) {
            $title = 'Chart';
        }
        $workbook->getProperties()->setCreator($presentation->getDocumentProperties()->getCreator())->setLastModifiedBy($presentation->getDocumentProperties()->getLastModifiedBy())->setTitle($title);

        // Add chart data
        $sheet = $workbook->setActiveSheetIndex(0);
        $sheet->setTitle('Sheet1');

        // Write series
        $seriesIndex = 0;
        foreach ($chart->getPlotArea()->getType()->getSeries() as $series) {
            // Title
            $sheet->setCellValueByColumnAndRow(1 + $seriesIndex, 1, $series->getTitle());

            // X-axis
            $axisXData = array_keys($series->getValues());
            $numAxisXData = count($axisXData);
            for ($i = 0; $i < $numAxisXData; $i++) {
                $sheet->setCellValueByColumnAndRow(0, $i + 2, $axisXData[$i]);
            }

            // Y-axis
            $axisYData = array_values($series->getValues());
            $numAxisYData = count($axisYData);
            for ($i = 0; $i < $numAxisYData; $i++) {
                $sheet->setCellValueByColumnAndRow(1 + $seriesIndex, $i + 2, $axisYData[$i]);
            }

            ++$seriesIndex;
        }

        // Save to string
        $writer = \PHPExcel_IOFactory::createWriter($workbook, 'Excel2007');
        $writer->save($tempName);

        // Load file in memory
        $returnValue = file_get_contents($tempName);
        if (@unlink($tempName) === false) {
            throw new \Exception('The file ' . $tempName . ' could not removed.');
        }

        return $returnValue;
    }

    /**
     * Write element with value attribute
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param string $elementName
     * @param string $value
     */
    protected function writeElementWithValAttribute($objWriter, $elementName, $value)
    {
        $objWriter->startElement($elementName);
        $objWriter->writeAttribute('val', $value);
        $objWriter->endElement();
    }

    /**
     * Write single value or reference
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param boolean $isReference
     * @param mixed $value
     * @param string $reference
     */
    protected function writeSingleValueOrReference($objWriter, $isReference, $value, $reference)
    {
        if (!$isReference) {
            // Value
            $objWriter->writeElement('c:v', $value);
            return;
        }

        // Reference and cache
        $objWriter->startElement('c:strRef');
        $objWriter->writeElement('c:f', $reference);
        $objWriter->startElement('c:strCache');
        $objWriter->startElement('c:ptCount');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        $objWriter->startElement('c:pt');
        $objWriter->writeAttribute('idx', '0');
        $objWriter->writeElement('c:v', $value);
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Write series value or reference
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param boolean $isReference
     * @param mixed $values
     * @param string $reference
     */
    protected function writeMultipleValuesOrReference($objWriter, $isReference, $values, $reference)
    {
        // c:strLit / c:numLit
        // c:strRef / c:numRef
        $referenceType = ($isReference ? 'Ref' : 'Lit');
        $dataType = 'str';
        if (is_int($values[0]) || is_float($values[0])) {
            $dataType = 'num';
        }
        $objWriter->startElement('c:' . $dataType . $referenceType);

        $numValues = count($values);
        if (!$isReference) {
            // Value

            // c:ptCount
            $objWriter->startElement('c:ptCount');
            $objWriter->writeAttribute('val', count($values));
            $objWriter->endElement();

            // Add points
            for ($i = 0; $i < $numValues; $i++) {
                // c:pt
                $objWriter->startElement('c:pt');
                $objWriter->writeAttribute('idx', $i);
                $objWriter->writeElement('c:v', $values[$i]);
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
            for ($i = 0; $i < $numValues; $i++) {
                // c:pt
                $objWriter->startElement('c:pt');
                $objWriter->writeAttribute('idx', $i);
                $objWriter->writeElement('c:v', $values[$i]);
                $objWriter->endElement();
            }

            $objWriter->endElement();
        }

        $objWriter->endElement();
    }

    /**
     * Write Title
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Title $subject
     * @throws \Exception
     */
    protected function writeTitle(XMLWriter $objWriter, Title $subject)
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
        $objWriter->writeAttribute('indent', CommonDrawing::pixelsToEmu($subject->getAlignment()->getIndent()));
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
        $objWriter->writeAttribute('strike', ($subject->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
        $objWriter->writeAttribute('sz', ($subject->getFont()->getSize() * 100));
        $objWriter->writeAttribute('u', $subject->getFont()->getUnderline());
        $objWriter->writeAttributeIf($subject->getFont()->isSuperScript(), 'baseline', '30000');
        $objWriter->writeAttributeIf($subject->getFont()->isSubScript(), 'baseline', '-25000');

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
     * Write Plot Area
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\PlotArea $subject
     * @param  \PhpOffice\PhpPresentation\Shape\Chart $chart
     * @throws \Exception
     */
    protected function writePlotArea(XMLWriter $objWriter, PlotArea $subject, Chart $chart)
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
        } elseif ($chartType instanceof Pie) {
            $this->writeTypePie($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Pie3D) {
            $this->writeTypePie3D($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Line) {
            $this->writeTypeLine($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } elseif ($chartType instanceof Scatter) {
            $this->writeTypeScatter($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
        } else {
            throw new \Exception('The chart type provided could not be rendered.');
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

    /**
     * Write Legend
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Legend $subject
     * @throws \Exception
     */
    protected function writeLegend(XMLWriter $objWriter, Legend $subject)
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
        if ($subject->getBorder()->getLineStyle() != Border::LINE_NONE) {
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
        $objWriter->writeAttribute('indent', CommonDrawing::pixelsToEmu($subject->getAlignment()->getIndent()));
        $objWriter->writeAttribute('lvl', $subject->getAlignment()->getLevel());

        // a:defRPr
        $objWriter->startElement('a:defRPr');

        $objWriter->writeAttribute('b', ($subject->getFont()->isBold() ? 'true' : 'false'));
        $objWriter->writeAttribute('i', ($subject->getFont()->isItalic() ? 'true' : 'false'));
        $objWriter->writeAttribute('strike', ($subject->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
        $objWriter->writeAttribute('sz', ($subject->getFont()->getSize() * 100));
        $objWriter->writeAttribute('u', $subject->getFont()->getUnderline());
        $objWriter->writeAttributeIf($subject->getFont()->isSuperScript(), 'baseline', '30000');
        $objWriter->writeAttributeIf($subject->getFont()->isSubScript(), 'baseline', '-25000');

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
     * Write Layout
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  mixed $subject
     * @throws \Exception
     */
    protected function writeLayout(XMLWriter $objWriter, $subject)
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

        if ($subject->getOffsetX() != 0) {
            // c:x
            $objWriter->startElement('c:x');
            $objWriter->writeAttribute('val', $subject->getOffsetX());
            $objWriter->endElement();
        }

        if ($subject->getOffsetY() != 0) {
            // c:y
            $objWriter->startElement('c:y');
            $objWriter->writeAttribute('val', $subject->getOffsetY());
            $objWriter->endElement();
        }

        if ($subject->getWidth() != 0) {
            // c:w
            $objWriter->startElement('c:w');
            $objWriter->writeAttribute('val', $subject->getWidth());
            $objWriter->endElement();
        }

        if ($subject->getHeight() != 0) {
            // c:h
            $objWriter->startElement('c:h');
            $objWriter->writeAttribute('val', $subject->getHeight());
            $objWriter->endElement();
        }

        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Write Type Area
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Type\Area $subject
     * @param  boolean $includeSheet
     * @throws \Exception
     */
    protected function writeTypeArea(XMLWriter $objWriter, Area $subject, $includeSheet = false)
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
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

            if ($series->getFill()->getFillType() != Fill::FILL_NONE) {
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
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
     * Write Type Bar
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Type\Bar $subject
     * @param  boolean $includeSheet
     * @throws \Exception
     */
    protected function writeTypeBar(XMLWriter $objWriter, Bar $subject, $includeSheet = false)
    {
        // c:bar3DChart
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // Fills for points?
            $dataPointFills = $series->getDataPointFills();
            foreach ($dataPointFills as $key => $value) {
                // c:dPt
                $objWriter->startElement('c:dPt');

                // c:idx
                $this->writeElementWithValAttribute($objWriter, 'c:idx', $key);

                if ($value->getFillType() != Fill::FILL_NONE) {
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
            $objWriter->writeAttribute('strike', ($series->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->isSuperScript(), 'baseline', '30000');
            $objWriter->writeAttributeIf($series->getFont()->isSubScript(), 'baseline', '-25000');

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
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

            // c:separator
            $objWriter->writeElementIf($series->hasShowSeparator(), 'c:separator', 'val', $series->getSeparator());

            $objWriter->endElement();

            // c:spPr
            if ($series->getFill()->getFillType() != Fill::FILL_NONE) {
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:overlap
        $objWriter->startElement('c:overlap');
        if ($subject->getBarGrouping() == Bar::GROUPING_CLUSTERED) {
            $objWriter->writeAttribute('val', '0');
        } elseif ($subject->getBarGrouping() == Bar::GROUPING_STACKED || $subject->getBarGrouping() == Bar::GROUPING_PERCENTSTACKED) {
            $objWriter->writeAttribute('val', '100');
        }
        $objWriter->endElement();

        // c:gapWidth
        $objWriter->startElement('c:gapWidth');
        $objWriter->writeAttribute('val', $subject->getGapWidthPercent());
        $objWriter->endElement();

        // c:shape
        $objWriter->startElement('c:shape');
        $objWriter->writeAttribute('val', 'box');
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
     * Write Type Bar3D
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D $subject
     * @param  boolean $includeSheet
     * @throws \Exception
     */
    protected function writeTypeBar3D(XMLWriter $objWriter, Bar3D $subject, $includeSheet = false)
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // Fills for points?
            $dataPointFills = $series->getDataPointFills();
            foreach ($dataPointFills as $key => $value) {
                // c:dPt
                $objWriter->startElement('c:dPt');

                // c:idx
                $this->writeElementWithValAttribute($objWriter, 'c:idx', $key);

                if ($value->getFillType() != Fill::FILL_NONE) {
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
            $objWriter->writeAttribute('strike', ($series->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->isSuperScript(), 'baseline', '30000');
            $objWriter->writeAttributeIf($series->getFont()->isSubScript(), 'baseline', '-25000');

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
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

            // c:separator
            $objWriter->writeElementIf($series->hasShowSeparator(), 'c:separator', 'val', $series->getSeparator());

            $objWriter->endElement();

            // c:spPr
            if ($series->getFill()->getFillType() != Fill::FILL_NONE) {
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:gapWidth
        $objWriter->startElement('c:gapWidth');
        $objWriter->writeAttribute('val', $subject->getGapWidthPercent());
        $objWriter->endElement();

        // c:shape
        $objWriter->startElement('c:shape');
        $objWriter->writeAttribute('val', 'box');
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
     * Write Type Pie
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Type\Pie $subject
     * @param  boolean $includeSheet
     * @throws \Exception
     */
    protected function writeTypePie(XMLWriter $objWriter, Pie $subject, $includeSheet = false)
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // Fills for points?
            $dataPointFills = $series->getDataPointFills();
            foreach ($dataPointFills as $key => $value) {
                // c:dPt
                $objWriter->startElement('c:dPt');

                // c:idx
                $this->writeElementWithValAttribute($objWriter, 'c:idx', $key);

                // c:spPr
                $objWriter->startElement('c:spPr');

                // Write fill
                $this->writeFill($objWriter, $value);

                $objWriter->endElement();

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
            $objWriter->writeAttribute('strike', ($series->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->isSuperScript(), 'baseline', '30000');
            $objWriter->writeAttributeIf($series->getFont()->isSubScript(), 'baseline', '-25000');

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
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

            // c:separator
            $objWriter->writeElementIf($series->hasShowSeparator(), 'c:separator', 'val', $series->getSeparator());

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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        $objWriter->endElement();
    }

    /**
     * Write Type Pie3D
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D $subject
     * @param  boolean $includeSheet
     * @throws \Exception
     */
    protected function writeTypePie3D(XMLWriter $objWriter, Pie3D $subject, $includeSheet = false)
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
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

                // c:idx
                $this->writeElementWithValAttribute($objWriter, 'c:idx', $key);

                // c:spPr
                $objWriter->startElement('c:spPr');

                // Write fill
                $this->writeFill($objWriter, $value);

                $objWriter->endElement();

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
            $objWriter->writeAttribute('strike', ($series->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->isSuperScript(), 'baseline', '30000');
            $objWriter->writeAttributeIf($series->getFont()->isSubScript(), 'baseline', '-25000');

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
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

            // c:separator
            $objWriter->writeElementIf($series->hasShowSeparator(), 'c:separator', 'val', $series->getSeparator());

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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        $objWriter->endElement();
    }

    /**
     * Write Type Line
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Type\Line $subject
     * @param  boolean $includeSheet
     * @throws \Exception
     */
    protected function writeTypeLine(XMLWriter $objWriter, Line $subject, $includeSheet = false)
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
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
            $objWriter->writeAttribute('strike', ($series->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->isSuperScript(), 'baseline', '30000');
            $objWriter->writeAttributeIf($series->getFont()->isSubScript(), 'baseline', '-25000');

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
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

            // c:separator
            $objWriter->writeElementIf($series->hasShowSeparator(), 'c:separator', 'val', $series->getSeparator());

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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:marker
        $objWriter->startElement('c:marker');
        $objWriter->writeAttribute('val', '1');
        $objWriter->endElement();

        // c:smooth
        $objWriter->startElement('c:smooth');
        $objWriter->writeAttribute('val', '0');
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
     * Write Type Scatter
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter $subject
     * @param  boolean $includeSheet
     * @throws \Exception
     */
    protected function writeTypeScatter(XMLWriter $objWriter, Scatter $subject, $includeSheet = false)
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
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
            $objWriter->writeAttribute('strike', ($series->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
            $objWriter->writeAttribute('sz', ($series->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $series->getFont()->getUnderline());
            $objWriter->writeAttributeIf($series->getFont()->isSuperScript(), 'baseline', '30000');
            $objWriter->writeAttributeIf($series->getFont()->isSubScript(), 'baseline', '-25000');

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            $this->writeColor($objWriter, $series->getFont()->getColor());

            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
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

            // c:separator
            $objWriter->writeElementIf($series->hasShowSeparator(), 'c:separator', 'val', $series->getSeparator());

            $objWriter->endElement();

            // c:spPr
            $objWriter->startElement('c:spPr');
            // Write fill
            $this->writeFill($objWriter, $series->getFill());
            // Write outline
            $this->writeOutline($objWriter, $series->getOutline());
            // ## c:spPr
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
            $coords = ($includeSheet ? 'Sheet1!$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . \PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();

            // c:smooth
            $objWriter->startElement('c:smooth');
            $objWriter->writeAttribute('val', '0');
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
     * Write chart relationships to XML format
     *
     * @param  \PhpOffice\PhpPresentation\Shape\Chart $pChart
     * @return string                    XML Output
     * @throws \Exception
     */
    public function writeChartRelationships(Chart $pChart)
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

    /**
     * @param XMLWriter $objWriter
     * @param Chart\Marker $oMarker
     */
    protected function writeSeriesMarker(XMLWriter $objWriter, Chart\Marker $oMarker)
    {
        // c:marker
        $objWriter->startElement('c:marker');
        // c:marker > c:symbol
        $objWriter->startElement('c:symbol');
        $objWriter->writeAttribute('val', $oMarker->getSymbol());
        $objWriter->endElement();

        // Size if different of none
        if ($oMarker->getSymbol() != Chart\Marker::SYMBOL_NONE) {
            $markerSize = (int)$oMarker->getSize();
            if ($markerSize < 2) {
                $markerSize = 2;
            }
            if ($markerSize > 72) {
                $markerSize = 72;
            }

            /**
             * c:marker > c:size
             * Size in points
             * @link : https://msdn.microsoft.com/en-us/library/hh658135(v=office.12).aspx
             */
            $objWriter->startElement('c:size');
            $objWriter->writeAttribute('val', $markerSize);
            $objWriter->endElement();
        }
        $objWriter->endElement();
    }

    /**
     * @param XMLWriter $objWriter
     * @param Chart\Axis $oAxis
     * @param $typeAxis
     * @param Chart\Type\AbstractType $typeChart
     */
    protected function writeAxis(XMLWriter $objWriter, Chart\Axis $oAxis, $typeAxis, Chart\Type\AbstractType $typeChart)
    {
        if ($typeAxis != Chart\Axis::AXIS_X && $typeAxis != Chart\Axis::AXIS_Y) {
            return;
        }

        if ($typeAxis == Chart\Axis::AXIS_X) {
            $mainElement = 'c:catAx';
            $axIdVal = '52743552';
            $axPosVal = 'b';
            $crossAxVal = '52749440';
        } else {
            $mainElement = 'c:valAx';
            $axIdVal = '52749440';
            $axPosVal = 'l';
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
        $objWriter->writeAttribute('val', 'minMax');
        $objWriter->endElement();

        if ($oAxis->getMaxBounds() != null) {
            $objWriter->startElement('c:max');
            $objWriter->writeAttribute('val', $oAxis->getMaxBounds());
            $objWriter->endElement();
        }

        if ($oAxis->getMinBounds() != null) {
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

        if ($oAxis->getTitle() != '') {
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

            // a:defRPr
            $objWriter->startElement('a:defRPr');

            $objWriter->writeAttribute('b', ($oAxis->getFont()->isBold() ? 'true' : 'false'));
            $objWriter->writeAttribute('i', ($oAxis->getFont()->isItalic() ? 'true' : 'false'));
            $objWriter->writeAttribute('strike', ($oAxis->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
            $objWriter->writeAttribute('sz', ($oAxis->getFont()->getSize() * 100));
            $objWriter->writeAttribute('u', $oAxis->getFont()->getUnderline());
            $objWriter->writeAttributeIf($oAxis->getFont()->isSuperScript(), 'baseline', '30000');
            $objWriter->writeAttributeIf($oAxis->getFont()->isSubScript(), 'baseline', '-25000');

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');
            $this->writeColor($objWriter, $oAxis->getFont()->getColor());
            $objWriter->endElement();

            // Font - a:latin
            $objWriter->startElement('a:latin');
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
        $objWriter->writeAttribute('val', 'nextTo');
        $objWriter->endElement();

        // c:spPr
        $objWriter->startElement('c:spPr');
        // Outline
        $this->writeOutline($objWriter, $oAxis->getOutline());
        // ##c:spPr
        $objWriter->endElement();

        // c:crossAx
        $objWriter->startElement('c:crossAx');
        $objWriter->writeAttribute('val', $crossAxVal);
        $objWriter->endElement();

        // c:crosses
        $objWriter->startElement('c:crosses');
        $objWriter->writeAttribute('val', 'autoZero');
        $objWriter->endElement();

        if ($typeAxis == Chart\Axis::AXIS_X) {
            // c:lblAlgn
            $objWriter->startElement('c:lblAlgn');
            $objWriter->writeAttribute('val', 'ctr');
            $objWriter->endElement();

            // c:lblOffset
            $objWriter->startElement('c:lblOffset');
            $objWriter->writeAttribute('val', '100%');
            $objWriter->endElement();
        }

        if ($typeAxis == Chart\Axis::AXIS_Y) {
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
            if ($oAxis->getMajorUnit() != null) {
                $objWriter->startElement('c:majorUnit');
                $objWriter->writeAttribute('val', $oAxis->getMajorUnit());
                $objWriter->endElement();
            }

            // c:minorUnit
            if ($oAxis->getMinorUnit() != null) {
                $objWriter->startElement('c:minorUnit');
                $objWriter->writeAttribute('val', $oAxis->getMinorUnit());
                $objWriter->endElement();
            }
        }

        $objWriter->endElement();
    }

    /**
     * @param XMLWriter $objWriter
     * @param Gridlines $oGridlines
     */
    protected function writeAxisGridlines(XMLWriter $objWriter, Gridlines $oGridlines)
    {
        // c:spPr
        $objWriter->startElement('c:spPr');

        // Outline
        $this->writeOutline($objWriter, $oGridlines->getOutline());

        // ##c:spPr
        $objWriter->endElement();
    }
}
