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
use PhpOffice\PhpPowerpoint\Shape\Chart\Legend;
use PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea;
use PhpOffice\PhpPowerpoint\Shape\Chart\Title;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Line;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPowerpoint\Shape\Chart as ShapeChart;
use PhpOffice\PhpPowerpoint\Shared\Drawing as SharedDrawing;
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;
use PhpOffice\PhpPowerpoint\Style\Border;

/**
 * \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Chart
 */
class Chart extends Slide
{
    /**
     * Write chart to XML format
     *
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart $chart
     * @return string                    XML Output
     * @throws \Exception
     */
    public function writeChart(ShapeChart $chart = null)
    {
        // Check slide
        if (is_null($chart)) {
            throw new \Exception("Invalid \PhpOffice\PhpPowerpoint\Shape\Chart object passed.");
        }

        // Create XML writer
        $objWriter = $this->getXMLWriter();

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
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        // c:view3D
        $objWriter->startElement('c:view3D');

        // c:rotX
        $objWriter->startElement('c:rotX');
        $objWriter->writeAttribute('val', $chart->getView3D()->getRotationX());
        $objWriter->endElement();

        // c:hPercent
        $objWriter->startElement('c:hPercent');
        $objWriter->writeAttribute('val', $chart->getView3D()->getHeightPercent());
        $objWriter->endElement();

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
            $objWriter->writeAttribute('blurRad', SharedDrawing::pixelsToEmu($chart->getShadow()->getBlurRadius()));
            $objWriter->writeAttribute('dist', SharedDrawing::pixelsToEmu($chart->getShadow()->getDistance()));
            $objWriter->writeAttribute('dir', SharedDrawing::degreesToAngle($chart->getShadow()->getDirection()));
            $objWriter->writeAttribute('algn', $chart->getShadow()->getAlignment());
            $objWriter->writeAttribute('rotWithShape', '0');

            // a:srgbClr
            $objWriter->startElement('a:srgbClr');
            $objWriter->writeAttribute('val', $chart->getShadow()->getColor()->getRGB());

            // a:alpha
            $objWriter->startElement('a:alpha');
            $objWriter->writeAttribute('val', $chart->getShadow()->getAlpha() * 1000);
            $objWriter->endElement();

            $objWriter->endElement();

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
     * @param  PHPPowerPoint             $presentation
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart $chart
     * @param  string                    $tempName
     * @return string                    String output
     * @throws \Exception
     */
    public function writeSpreadsheet(PHPPowerPoint $presentation, $chart, $tempName)
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
        $workbook = new PHPExcel();

        // Set properties
        $title = $chart->getTitle()->getText();
        if (strlen($title) == 0) {
            $title = 'Chart';
        }
        $workbook->getProperties()->setCreator($presentation->getProperties()->getCreator())->setLastModifiedBy($presentation->getProperties()->getLastModifiedBy())->setTitle($title);

        // Add chart data
        $sheet = $workbook->setActiveSheetIndex(0);
        $sheet->setTitle('Sheet1');

        // Write series
        $seriesIndex = 0;
        foreach ($chart->getPlotArea()->getType()->getData() as $series) {
            // Title
            $sheet->setCellValueByColumnAndRow(1 + $seriesIndex, 1, $series->getTitle());

            // X-axis
            $axisXData = array_keys($series->getValues());
            for ($i = 0; $i < count($axisXData); $i++) {
                $sheet->setCellValueByColumnAndRow(0, $i + 2, $axisXData[$i]);
            }

            // Y-axis
            $axisYData = array_values($series->getValues());
            for ($i = 0; $i < count($axisYData); $i++) {
                $sheet->setCellValueByColumnAndRow(1 + $seriesIndex, $i + 2, $axisYData[$i]);
            }

            ++$seriesIndex;
        }

        // Save to string
        $writer = PHPExcel_IOFactory::createWriter($workbook, 'Excel2007');
        $writer->save($tempName);

        // Load file in memory
        $returnValue = file_get_contents($tempName);
        @unlink($tempName);

        return $returnValue;
    }

    /**
     * Write element with value attribute
     *
     * @param \PhpOffice\PhpPowerpoint\Shared\XMLWriter $objWriter   XML Writer
     * @param string                         $elementName
     * @param string                         $value
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
     * @param \PhpOffice\PhpPowerpoint\Shared\XMLWriter $objWriter   XML Writer
     * @param boolean                        $isReference
     * @param mixed                          $value
     * @param string                         $reference
     */
    protected function writeSingleValueOrReference($objWriter, $isReference, $value, $reference)
    {
        if (!$isReference) {
            // Value
            $objWriter->writeElement('c:v', $value);
        } else {
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
    }

    /**
     * Write series value or reference
     *
     * @param \PhpOffice\PhpPowerpoint\Shared\XMLWriter $objWriter   XML Writer
     * @param boolean                        $isReference
     * @param mixed                          $values
     * @param string                         $reference
     */
    protected function writeMultipleValuesOrReference($objWriter, $isReference, $values, $reference)
    {
        // c:strLit / c:numLit
        // c:strRef / c:numRef
        $referenceType = ($isReference ? 'Ref' : 'Lit');
        if (is_int($values[0]) || is_float($values[0])) {
            $dataType = 'num';
        } else {
            $dataType = 'str';
        }
        $objWriter->startElement('c:' . $dataType . $referenceType);

        if (!$isReference) {
            // Value

            // c:ptCount
            $objWriter->startElement('c:ptCount');
            $objWriter->writeAttribute('val', count($values));
            $objWriter->endElement();

            // Add points
            for ($i = 0; $i < count($values); $i++) {
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
            for ($i = 0; $i < count($values); $i++) {
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
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter  $objWriter XML Writer
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart\Title $subject
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
        $objWriter->writeAttribute('marL', SharedDrawing::pixelsToEmu($subject->getAlignment()->getMarginLeft()));
        $objWriter->writeAttribute('marR', SharedDrawing::pixelsToEmu($subject->getAlignment()->getMarginRight()));
        $objWriter->writeAttribute('indent', SharedDrawing::pixelsToEmu($subject->getAlignment()->getIndent()));
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

        if ($subject->getFont()->isSuperScript() || $subject->getFont()->isSubScript()) {
            if ($subject->getFont()->isSuperScript()) {
                $objWriter->writeAttribute('baseline', '30000');
            } elseif ($subject->getFont()->isSubScript()) {
                $objWriter->writeAttribute('baseline', '-25000');
            }
        }

        // Font - a:solidFill
        $objWriter->startElement('a:solidFill');

        // a:srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $subject->getFont()->getColor()->getRGB());
        $objWriter->endElement();

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
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter     $objWriter XML Writer
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea $subject
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart          $chart
     * @throws \Exception
     */
    protected function writePlotArea(XMLWriter $objWriter, PlotArea $subject, ShapeChart $chart)
    {
        // c:plotArea
        $objWriter->startElement('c:plotArea');

        // Write layout
        $this->writeLayout($objWriter, $subject);

        // Write chart
        $chartType = $subject->getType();
        if ($chartType instanceof Bar3D) {
            $this->writeTypeBar3D($objWriter, $chartType, $chart->hasIncludedSpreadsheet());
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
            // c:catAx (Axis X)
            $objWriter->startElement('c:catAx');

            // c:axId
            $objWriter->startElement('c:axId');
            $objWriter->writeAttribute('val', '52743552');
            $objWriter->endElement();

            // c:scaling
            $objWriter->startElement('c:scaling');

            // c:orientation
            $objWriter->startElement('c:orientation');
            $objWriter->writeAttribute('val', 'minMax');
            $objWriter->endElement();

            $objWriter->endElement();

            // c:axPos
            $objWriter->startElement('c:axPos');
            $objWriter->writeAttribute('val', 'b');
            $objWriter->endElement();

            // c:numFmt
            $objWriter->startElement('c:numFmt');
            $objWriter->writeAttribute('formatCode', $subject->getAxisX()->getFormatCode());
            $objWriter->writeAttribute('sourceLinked', '0');
            $objWriter->endElement();

            // c:majorTickMark
            $objWriter->startElement('c:majorTickMark');
            $objWriter->writeAttribute('val', 'none');
            $objWriter->endElement();

            // c:tickLblPos
            $objWriter->startElement('c:tickLblPos');
            $objWriter->writeAttribute('val', 'nextTo');
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
            $objWriter->writeElement('a:defRPr', null);

            $objWriter->endElement();

            // a:r
            $objWriter->startElement('a:r');

            // a:rPr
            $objWriter->startElement('a:rPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            // a:t
            $objWriter->writeElement('a:t', $subject->getAxisX()->getTitle());

            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // c:crossAx
            $objWriter->startElement('c:crossAx');
            $objWriter->writeAttribute('val', '52749440');
            $objWriter->endElement();

            // c:crosses
            $objWriter->startElement('c:crosses');
            $objWriter->writeAttribute('val', 'autoZero');
            $objWriter->endElement();

            // c:lblAlgn
            $objWriter->startElement('c:lblAlgn');
            $objWriter->writeAttribute('val', 'ctr');
            $objWriter->endElement();

            // c:lblOffset
            $objWriter->startElement('c:lblOffset');
            $objWriter->writeAttribute('val', '100');
            $objWriter->endElement();

            $objWriter->endElement();
        }

        // Write Y axis?
        if ($chartType->hasAxisY()) {
            // c:valAx (Axis Y)
            $objWriter->startElement('c:valAx');

            // c:axId
            $objWriter->startElement('c:axId');
            $objWriter->writeAttribute('val', '52749440');
            $objWriter->endElement();

            // c:scaling
            $objWriter->startElement('c:scaling');

            // c:orientation
            $objWriter->startElement('c:orientation');
            $objWriter->writeAttribute('val', 'minMax');
            $objWriter->endElement();

            $objWriter->endElement();

            // c:axPos
            $objWriter->startElement('c:axPos');
            $objWriter->writeAttribute('val', 'l');
            $objWriter->endElement();

            // c:numFmt
            $objWriter->startElement('c:numFmt');
            $objWriter->writeAttribute('formatCode', $subject->getAxisY()->getFormatCode());
            $objWriter->writeAttribute('sourceLinked', '0');
            $objWriter->endElement();

            // c:majorGridlines
            //$objWriter->startElement('c:majorGridlines');
            //$objWriter->endElement();

            // c:majorTickMark
            $objWriter->startElement('c:majorTickMark');
            $objWriter->writeAttribute('val', 'none');
            $objWriter->endElement();

            // c:tickLblPos
            $objWriter->startElement('c:tickLblPos');
            $objWriter->writeAttribute('val', 'nextTo');
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
            $objWriter->writeElement('a:defRPr', null);

            $objWriter->endElement();

            // a:r
            $objWriter->startElement('a:r');

            // a:rPr
            $objWriter->startElement('a:rPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            // a:t
            $objWriter->writeElement('a:t', $subject->getAxisY()->getTitle());

            $objWriter->endElement();

            // a:endParaRPr
            $objWriter->startElement('a:endParaRPr');
            $objWriter->writeAttribute('lang', 'en-US');
            $objWriter->writeAttribute('dirty', '0');
            $objWriter->endElement();

            $objWriter->endElement();

            $objWriter->endElement();

            // c:crossAx
            $objWriter->startElement('c:crossAx');
            $objWriter->writeAttribute('val', '52743552');
            $objWriter->endElement();

            // c:crosses
            $objWriter->startElement('c:crosses');
            $objWriter->writeAttribute('val', 'autoZero');
            $objWriter->endElement();

            // c:crossBetween
            $objWriter->startElement('c:crossBetween');
            $objWriter->writeAttribute('val', 'between');
            $objWriter->endElement();

            $objWriter->endElement();
        }

        $objWriter->endElement();
    }

    /**
     * Write Legend
     *
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter   $objWriter XML Writer
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart\Legend $subject
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
        $objWriter->writeAttribute('marL', SharedDrawing::pixelsToEmu($subject->getAlignment()->getMarginLeft()));
        $objWriter->writeAttribute('marR', SharedDrawing::pixelsToEmu($subject->getAlignment()->getMarginRight()));
        $objWriter->writeAttribute('indent', SharedDrawing::pixelsToEmu($subject->getAlignment()->getIndent()));
        $objWriter->writeAttribute('lvl', $subject->getAlignment()->getLevel());

        // a:defRPr
        $objWriter->startElement('a:defRPr');

        $objWriter->writeAttribute('b', ($subject->getFont()->isBold() ? 'true' : 'false'));
        $objWriter->writeAttribute('i', ($subject->getFont()->isItalic() ? 'true' : 'false'));
        $objWriter->writeAttribute('strike', ($subject->getFont()->isStrikethrough() ? 'sngStrike' : 'noStrike'));
        $objWriter->writeAttribute('sz', ($subject->getFont()->getSize() * 100));
        $objWriter->writeAttribute('u', $subject->getFont()->getUnderline());

        if ($subject->getFont()->isSuperScript() || $subject->getFont()->isSubScript()) {
            if ($subject->getFont()->isSuperScript()) {
                $objWriter->writeAttribute('baseline', '30000');
            } elseif ($subject->getFont()->isSubScript()) {
                $objWriter->writeAttribute('baseline', '-25000');
            }
        }

        // Font - a:solidFill
        $objWriter->startElement('a:solidFill');

        // a:srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $subject->getFont()->getColor()->getRGB());
        $objWriter->endElement();

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
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter $objWriter XML Writer
     * @param  mixed                          $subject
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
     * Write Type Bar3D
     *
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter       $objWriter    XML Writer
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D $subject
     * @param  boolean                              $includeSheet
     * @throws \Exception
     */
    protected function writeTypeBar3D(XMLWriter $objWriter, Bar3D $subject, $includeSheet = false)
    {
        // c:bar3DChart
        $objWriter->startElement('c:bar3DChart');

        // c:barDir
        $objWriter->startElement('c:barDir');
        $objWriter->writeAttribute('val', 'col');
        $objWriter->endElement();

        // c:grouping
        $objWriter->startElement('c:grouping');
        $objWriter->writeAttribute('val', 'clustered');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getData() as $series) {
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
            $coords = ($includeSheet ? 'Sheet1!$' . PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
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

            if ($series->getFont()->isSuperScript() || $series->getFont()->isSubScript()) {
                if ($series->getFont()->isSuperScript()) {
                    $objWriter->writeAttribute('baseline', '30000');
                } elseif ($series->getFont()->isSubScript()) {
                    $objWriter->writeAttribute('baseline', '-25000');
                }
            }

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            // a:srgbClr
            $objWriter->startElement('a:srgbClr');
            $objWriter->writeAttribute('val', $series->getFont()->getColor()->getRGB());
            $objWriter->endElement();

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

            $objWriter->endElement();

            // c:spPr
            $objWriter->startElement('c:spPr');

            // Write fill
            $this->writeFill($objWriter, $series->getFill());

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
            $coords = ($includeSheet ? 'Sheet1!$' . PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
            $this->writeMultipleValuesOrReference($objWriter, $includeSheet, $axisYData, $coords);
            $objWriter->endElement();
            
            $objWriter->endElement();

            ++$seriesIndex;
        }

        // c:gapWidth
        $objWriter->startElement('c:gapWidth');
        $objWriter->writeAttribute('val', '75');
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
     * Write Type Pie3D
     *
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter       $objWriter    XML Writer
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart\Type\Pie3D $subject
     * @param  boolean                              $includeSheet
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
        foreach ($subject->getData() as $series) {
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
            $coords = ($includeSheet ? 'Sheet1!$' . PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // c:explosion
            $objWriter->startElement('c:explosion');
            $objWriter->writeAttribute('val', '20');
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

            if ($series->getFont()->isSuperScript() || $series->getFont()->isSubScript()) {
                if ($series->getFont()->isSuperScript()) {
                    $objWriter->writeAttribute('baseline', '30000');
                } elseif ($series->getFont()->isSubScript()) {
                    $objWriter->writeAttribute('baseline', '-25000');
                }
            }

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            // a:srgbClr
            $objWriter->startElement('a:srgbClr');
            $objWriter->writeAttribute('val', $series->getFont()->getColor()->getRGB());
            $objWriter->endElement();

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
            $coords = ($includeSheet ? 'Sheet1!$' . PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
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
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter      $objWriter    XML Writer
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart\Type\Line $subject
     * @param  boolean                             $includeSheet
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

        // c:varyColors
        $objWriter->startElement('c:varyColors');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        // Write series
        $seriesIndex = 0;
        foreach ($subject->getData() as $series) {
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
            $coords = ($includeSheet ? 'Sheet1!$' . PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

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

            if ($series->getFont()->isSuperScript() || $series->getFont()->isSubScript()) {
                if ($series->getFont()->isSuperScript()) {
                    $objWriter->writeAttribute('baseline', '30000');
                } elseif ($series->getFont()->isSubScript()) {
                    $objWriter->writeAttribute('baseline', '-25000');
                }
            }

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            // a:srgbClr
            $objWriter->startElement('a:srgbClr');
            $objWriter->writeAttribute('val', $series->getFont()->getColor()->getRGB());
            $objWriter->endElement();

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

            $objWriter->endElement();

            // c:spPr
            $objWriter->startElement('c:spPr');

            // Write fill
            $this->writeFill($objWriter, $series->getFill());

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
            $coords = ($includeSheet ? 'Sheet1!$' . PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
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

        // c:axId
        $objWriter->startElement('c:axId');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Type Scatter
     *
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter         $objWriter    XML Writer
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart\Type\Scatter $subject
     * @param  boolean                                $includeSheet
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
        foreach ($subject->getData() as $series) {
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
            $coords = ($includeSheet ? 'Sheet1!$' . PHPExcel_Cell::stringFromColumnIndex(1 + $seriesIndex) . '$1' : '');
            $this->writeSingleValueOrReference($objWriter, $includeSheet, $series->getTitle(), $coords);
            $objWriter->endElement();

            // c:marker
            $objWriter->startElement('c:marker');

            // c:marker
            $objWriter->startElement('c:symbol');
            $objWriter->writeAttribute('val', 'none'); // Marker style
            //$objWriter->writeAttribute('size', '7'); // Marker size
            $objWriter->endElement();

            $objWriter->endElement();
            
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

            if ($series->getFont()->isSuperScript() || $series->getFont()->isSubScript()) {
                if ($series->getFont()->isSuperScript()) {
                    $objWriter->writeAttribute('baseline', '30000');
                } elseif ($series->getFont()->isSubScript()) {
                    $objWriter->writeAttribute('baseline', '-25000');
                }
            }

            // Font - a:solidFill
            $objWriter->startElement('a:solidFill');

            // a:srgbClr
            $objWriter->startElement('a:srgbClr');
            $objWriter->writeAttribute('val', $series->getFont()->getColor()->getRGB());
            $objWriter->endElement();

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
            $this->writeElementWithValAttribute($objWriter, 'c:showLegendKey', '0');

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
            /*
            // c:spPr
            $objWriter->startElement('c:spPr');

            // Write fill
            $this->writeFill($objWriter, $series->getFill());

            $objWriter->endElement();
            */
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
            $coords = ($includeSheet ? 'Sheet1!$' . PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$2:$' . PHPExcel_Cell::stringFromColumnIndex($seriesIndex + 1) . '$' . (1 + count($axisYData)) : '');
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
}
