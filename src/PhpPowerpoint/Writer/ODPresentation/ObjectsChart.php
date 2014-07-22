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

namespace PhpOffice\PhpPowerpoint\Writer\ODPresentation;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\Chart;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Line;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPowerpoint\Shared\Drawing as SharedDrawing;
use PhpOffice\PhpPowerpoint\Shared\String;
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;
use PhpOffice\PhpPowerpoint\Style\Fill;

/**
 * \PhpOffice\PhpPowerpoint\Writer\ODPresentation\Objects
 */
class ObjectsChart extends AbstractPart
{
    /**
     * @var XMLWriter
     */
    private $xmlContent;
    /**
     * @var XMLWriter
     */
    private $xmlMeta;
    /**
     * @var XMLWriter
     */
    private $xmlStyles;
    /**
     * @var mixed
     */
    private $arrayData;
    /**
     * @var mixed
     */
    private $arrayTitle;
    /**
     * @var integer
     */
    private $numData;
    /**
     * @var integer
     */
    private $numSeries;
    /**
     * @var string
     */
    private $rangeCol;
    
    public function writePart(Chart $chart)
    {
        $this->xmlContent = $this->getXMLWriter();
        $this->xmlMeta = $this->getXMLWriter();
        $this->xmlStyles = $this->getXMLWriter();
        
        $this->writeContentPart($chart);
        
        return array(
            'content.xml' => $this->xmlContent->getData(),
            'meta.xml' => $this->xmlMeta->getData(),
            'styles.xml' => $this->xmlStyles->getData(),
        );
    }
    
    /**
     * @param Chart $chart
     */
    private function writeContentPart(Chart $chart)
    {
        $chartType = $chart->getPlotArea()->getType();
        if (!($chartType instanceof Bar3D || $chartType instanceof Line || $chartType instanceof Pie3D|| $chartType instanceof Scatter)) {
            throw new \Exception('The chart type provided could not be rendered.');
        }
        
        // Data
        $this->arrayData = array();
        $this->arrayTitle = array();
        $this->numData = 0;
        foreach ($chart->getPlotArea()->getType()->getData() as $series) {
            $inc = 0;
            $this->arrayTitle[] = $series->getTitle();
            foreach ($series->getValues() as $key => $value) {
                if (!isset($this->arrayData[$inc])) {
                    $this->arrayData[$inc] = array();
                }
                if (empty($this->arrayData[$inc])) {
                    $this->arrayData[$inc][] = $key;
                }
                $this->arrayData[$inc][] = $value;
                $inc++;
            }
            if ($inc > $this->numData) {
                $this->numData = $inc;
            }
        }
        
        // office:document-content
        $this->xmlContent->startElement('office:document-content');
        $this->xmlContent->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
        $this->xmlContent->writeAttribute('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0');
        $this->xmlContent->writeAttribute('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
        $this->xmlContent->writeAttribute('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
        $this->xmlContent->writeAttribute('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
        $this->xmlContent->writeAttribute('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0');
        $this->xmlContent->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
        $this->xmlContent->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $this->xmlContent->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
        $this->xmlContent->writeAttribute('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0');
        $this->xmlContent->writeAttribute('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0');
        $this->xmlContent->writeAttribute('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0');
        $this->xmlContent->writeAttribute('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0');
        $this->xmlContent->writeAttribute('xmlns:math', 'http://www.w3.org/1998/Math/MathML');
        $this->xmlContent->writeAttribute('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0');
        $this->xmlContent->writeAttribute('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0');
        $this->xmlContent->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
        $this->xmlContent->writeAttribute('xmlns:ooow', 'http://openoffice.org/2004/writer');
        $this->xmlContent->writeAttribute('xmlns:oooc', 'http://openoffice.org/2004/calc');
        $this->xmlContent->writeAttribute('xmlns:dom', 'http://www.w3.org/2001/xml-events');
        $this->xmlContent->writeAttribute('xmlns:xforms', 'http://www.w3.org/2002/xforms');
        $this->xmlContent->writeAttribute('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema');
        $this->xmlContent->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
        $this->xmlContent->writeAttribute('xmlns:rpt', 'http://openoffice.org/2005/report');
        $this->xmlContent->writeAttribute('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2');
        $this->xmlContent->writeAttribute('xmlns:xhtml', 'http://www.w3.org/1999/xhtml');
        $this->xmlContent->writeAttribute('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#');
        $this->xmlContent->writeAttribute('xmlns:tableooo', 'http://openoffice.org/2009/table');
        $this->xmlContent->writeAttribute('xmlns:chartooo', 'http://openoffice.org/2010/chart');
        $this->xmlContent->writeAttribute('xmlns:drawooo', 'http://openoffice.org/2010/draw');
        $this->xmlContent->writeAttribute('xmlns:calcext', 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0');
        $this->xmlContent->writeAttribute('xmlns:loext', 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0');
        $this->xmlContent->writeAttribute('xmlns:field', 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0');
        $this->xmlContent->writeAttribute('xmlns:formx', 'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0');
        $this->xmlContent->writeAttribute('xmlns:css3t', 'http://www.w3.org/TR/css3-text/');
        $this->xmlContent->writeAttribute('office:version', '1.2');
        
        // office:automatic-styles
        $this->xmlContent->startElement('office:automatic-styles');
        
        // Chart
        $this->writeChartStyle($chart);
        
        // Axis
        $this->writeAxisStyle();
        
        // Series
        $this->numSeries = 0;
        foreach ($chart->getPlotArea()->getType()->getData() as $series) {
            $this->writeSeriesStyle($chart, $series);
            
            $this->numSeries++;
        }
        
        // Floor
        $this->writeFloorStyle();
        
        // Legend
        $this->writeLegendStyle($chart);
        
        // PlotArea
        $this->writePlotAreaStyle($chart);
        
        // Title
        $this->writeTitleStyle($chart);
        
        // Wall
        $this->writeWallStyle($chart);
        
        // > office:automatic-styles
        $this->xmlContent->endElement();
        
        // office:body
        $this->xmlContent->startElement('office:body');
        // office:chart
        $this->xmlContent->startElement('office:chart');
        // office:chart
        $this->xmlContent->startElement('chart:chart');
        $this->xmlContent->writeAttribute('svg:width', String::numberFormat(SharedDrawing::pixelsToCentimeters($chart->getWidth()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:height', String::numberFormat(SharedDrawing::pixelsToCentimeters($chart->getHeight()), 3) . 'cm');
        $this->xmlContent->writeAttribute('xlink:href', '.');
        $this->xmlContent->writeAttribute('xlink:type', 'simple');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleChart');
        if ($chartType instanceof Bar3D) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:bar');
        } elseif ($chartType instanceof Line) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:line');
        } elseif ($chartType instanceof Pie3D) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:circle');
        } elseif ($chartType instanceof Scatter) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:scatter');
        }
        
        //**** Title ****
        $this->writeTitle($chart);

        //**** Legend ****
        $this->writeLegend($chart);
        
        //**** Plotarea ****
        $this->writePlotArea($chart);

        //**** Table ****
        $this->writeTable();
        
        // > chart:chart
        $this->xmlContent->endElement();
        // > office:chart
        $this->xmlContent->endElement();
        // > office:body
        $this->xmlContent->endElement();
        // > office:document-content
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     */
    private function writeAxis(Chart $chart)
    {
        $chartType = $chart->getPlotArea()->getType();
        
        // chart:axis
        $this->xmlContent->startElement('chart:axis');
        $this->xmlContent->writeAttribute('chart:dimension', 'x');
        $this->xmlContent->writeAttribute('chart:name', 'primary-x');
        $this->xmlContent->writeAttribute('chartooo:axis-type', 'text');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleAxisX');
        // chart:categories
        $this->xmlContent->startElement('chart:categories');
        $this->xmlContent->writeAttribute('table:cell-range-address', 'table-local.$A$2:.$A$'.($this->numData+1));
        // > chart:categories
        $this->xmlContent->endElement();
        // > chart:axis
        $this->xmlContent->endElement();
        // chart:axis
        $this->xmlContent->startElement('chart:axis');
        $this->xmlContent->writeAttribute('chart:dimension', 'y');
        $this->xmlContent->writeAttribute('chart:name', 'primary-y');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleAxisY');
        // > chart:axis
        $this->xmlContent->endElement();
        if ($chartType instanceof Bar3D || $chartType instanceof Pie3D) {
            // chart:axis
            $this->xmlContent->startElement('chart:axis');
            $this->xmlContent->writeAttribute('chart:dimension', 'z');
            $this->xmlContent->writeAttribute('chart:name', 'primary-z');
            // > chart:axis
            $this->xmlContent->endElement();
        }
    }
    
    private function writeAxisStyle()
    {
        // AxisX
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleAxisX');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        $this->xmlContent->writeAttribute('chart:display-label', 'true');
        // > style:chart-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
        
        // AxisY
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleAxisY');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        $this->xmlContent->writeAttribute('chart:display-label', 'true');
        // > style:chart-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     */
    private function writeChartStyle(Chart $chart)
    {
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleChart');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        $this->xmlContent->writeAttribute('draw:stroke', $chart->getFill()->getFillType());
        $this->xmlContent->writeAttribute('draw:fill-color', '#'.$chart->getFill()->getStartColor()->getRGB());
        // > style:graphic-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }
    
    private function writeFloor()
    {
        // chart:floor
        $this->xmlContent->startElement('chart:floor');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleFloor');
        // > chart:floor
        $this->xmlContent->endElement();
    }
    
    private function writeFloorStyle()
    {
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleFloor');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:chart-properties
        $this->xmlContent->startElement('style:graphic-properties');
        $this->xmlContent->writeAttribute('draw:fill', 'none');
        //@todo : Permit edit color and size border of floor
        $this->xmlContent->writeAttribute('draw:stroke', 'solid');
        $this->xmlContent->writeAttribute('svg:stroke-width', '0.026cm');
        $this->xmlContent->writeAttribute('svg:stroke-color', '#878787');
        // > style:chart-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     */
    private function writeLegend(Chart $chart)
    {
        // chart:legend
        $this->xmlContent->startElement('chart:legend');
        $this->xmlContent->writeAttribute('chart:legend-position', 'end');
        $this->xmlContent->writeAttribute('svg:x', String::numberFormat(SharedDrawing::pixelsToCentimeters($chart->getLegend()->getOffsetX()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:y', String::numberFormat(SharedDrawing::pixelsToCentimeters($chart->getLegend()->getOffsetY()), 3) . 'cm');
        $this->xmlContent->writeAttribute('style:legend-expansion', 'high');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleLegend');
        // > chart:legend
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     */
    private function writeLegendStyle(Chart $chart)
    {
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleLegend');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:text-properties
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#'.$chart->getLegend()->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $chart->getLegend()->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $chart->getLegend()->getFont()->getSize().'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $chart->getLegend()->getFont()->isItalic() ? 'italic' : 'normal');
        // > style:text-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     */
    private function writePlotArea(Chart $chart)
    {
        $chartType = $chart->getPlotArea()->getType();
        
        // chart:plot-area
        $this->xmlContent->startElement('chart:plot-area');
        $this->xmlContent->writeAttribute('chart:style-name', 'stylePlotArea');
        if ($chartType instanceof Bar3D || $chartType instanceof Pie3D) {
            $this->xmlContent->writeAttribute('dr3d:ambient-color', '#cccccc');
            $this->xmlContent->writeAttribute('dr3d:lighting-mode', 'true');
        }
        if ($chartType instanceof Bar3D || $chartType instanceof Pie3D) {
            // dr3d:light
            $arrayLight = array(
                    array('#808080', '(0 0 1)', 'false', 'true'),
                    array('#666666', '(0.2 0.4 1)', 'true', 'false'),
                    array('#808080', '(0 0 1)', 'false', 'false'),
                    array('#808080', '(0 0 1)', 'false', 'false'),
                    array('#808080', '(0 0 1)', 'false', 'false'),
                    array('#808080', '(0 0 1)', 'false', 'false'),
                    array('#808080', '(0 0 1)', 'false', 'false'),
            );
            foreach ($arrayLight as $light) {
                $this->xmlContent->startElement('dr3d:light');
                $this->xmlContent->writeAttribute('dr3d:diffuse-color', $light[0]);
                $this->xmlContent->writeAttribute('dr3d:direction', $light[1]);
                $this->xmlContent->writeAttribute('dr3d:enabled', $light[2]);
                $this->xmlContent->writeAttribute('dr3d:specular', $light[3]);
                $this->xmlContent->endElement();
            }
        }
        
        //**** Axis ****
        $this->writeAxis($chart);
        
        //**** Series ****
        $this->rangeCol = 'B';
        $this->numSeries = 0;
        foreach ($chart->getPlotArea()->getType()->getData() as $series) {
            $this->writeSeries($chart, $series);
            $this->rangeCol++;
            $this->numSeries++;
        }
        
        //**** Wall ****
        $this->writeWall();
        //**** Floor ****
        $this->writeFloor();
        // > chart:plot-area
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     */
    private function writePlotAreaStyle(Chart $chart)
    {
        $chartType = $chart->getPlotArea()->getType();
        
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'stylePlotArea');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:text-properties
        $this->xmlContent->startElement('style:chart-properties');
        if ($chartType instanceof Bar3D || $chartType instanceof Pie3D) {
            $this->xmlContent->writeAttribute('chart:three-dimensional', 'true');
            $this->xmlContent->writeAttribute('chart:right-angled-axes', 'true');
        }
        $this->xmlContent->writeAttribute('chart:data-label-number', 'value');
        // > style:text-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     * @param Chart\Series $series
     * @throws \Exception
     */
    private function writeSeries(Chart $chart, Chart\Series $series)
    {
        $chartType = $chart->getPlotArea()->getType();
        
        $numRange = count($series->getValues());
        // chart:series
        $this->xmlContent->startElement('chart:series');
        $this->xmlContent->writeAttribute('chart:values-cell-range-address', 'table-local.$'.$this->rangeCol.'$2:.$'.$this->rangeCol.'$'.($numRange+1));
        $this->xmlContent->writeAttribute('chart:label-cell-address', 'table-local.$'.$this->rangeCol.'$1');
        if ($chartType instanceof Bar3D) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:bar');
        } elseif ($chartType instanceof Line) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:line');
        } elseif ($chartType instanceof Pie3D) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:circle');
        } elseif ($chartType instanceof Scatter) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:scatter');
        }
        $this->xmlContent->writeAttribute('chart:style-name', 'styleSeries'.$this->numSeries);
        if ($chartType instanceof Bar3D || $chartType instanceof Line || $chartType instanceof Scatter) {
            $dataPointFills = $series->getDataPointFills();
            if (empty($dataPointFills)) {
                $incRepeat = $numRange;
            } else {
                $inc = 0;
                $incRepeat = 0;
                $newFill = new Fill();
                do {
                    if ($series->getDataPointFill($inc)->getHashCode() != $newFill->getHashCode()) {
                        // chart:data-point
                        $this->xmlContent->startElement('chart:data-point');
                        $this->xmlContent->writeAttribute('chart:repeated', $incRepeat);
                        // > chart:data-point
                        $this->xmlContent->endElement();
                        $incRepeat = 0;
        
                        // chart:data-point
                        $this->xmlContent->startElement('chart:data-point');
                        $this->xmlContent->writeAttribute('chart:style-name', 'styleSeries'.$this->numSeries.'_'.$inc);
                        // > chart:data-point
                        $this->xmlContent->endElement();
                    }
                    $inc++;
                    $incRepeat++;
                } while ($inc < $numRange);
                $incRepeat--;
            }
            // chart:data-point
            $this->xmlContent->startElement('chart:data-point');
            $this->xmlContent->writeAttribute('chart:repeated', $incRepeat);
            // > chart:data-point
            $this->xmlContent->endElement();
        } elseif ($chartType instanceof Pie3D) {
            $count = count($series->getDataPointFills());
            for ($inc = 0; $inc < $count; $inc++) {
                // chart:data-point
                $this->xmlContent->startElement('chart:data-point');
                $this->xmlContent->writeAttribute('chart:style-name', 'styleSeries'.$this->numSeries.'_'.$inc);
                // > chart:data-point
                $this->xmlContent->endElement();
            }
        }
        
        // > chart:series
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     * @param Chart\Series $series
     */
    private function writeSeriesStyle(Chart $chart, Chart\Series $series)
    {
        $chartType = $chart->getPlotArea()->getType();
        
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleSeries'.$this->numSeries);
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        $this->xmlContent->writeAttribute('chart:data-label-number', 'value');
        $this->xmlContent->writeAttribute('chart:label-position', 'right');
        if ($chartType instanceof Pie3D) {
            //@todo : Permit edit the offset of a pie
            $this->xmlContent->writeAttribute('chart:pie-offset', '20');
        }
        if ($chartType instanceof Line) {
            //@todo : Permit edit the symbol of a line
            $this->xmlContent->writeAttribute('chart:symbol-type', 'automatic');
        }
        // > style:chart-properties
        $this->xmlContent->endElement();
        // style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        if ($chartType instanceof Line || $chartType instanceof Scatter) {
            //@todo : Permit edit the color and width of a line
            $this->xmlContent->writeAttribute('svg:stroke-width', '0.079cm');
            $this->xmlContent->writeAttribute('svg:stroke-color', '#4a7ebb');
            $this->xmlContent->writeAttribute('draw:fill-color', '#4a7ebb');
        } else {
            $this->xmlContent->writeAttribute('draw:stroke', 'none');
            $this->xmlContent->writeAttribute('draw:fill', $series->getFill()->getFillType());
            $this->xmlContent->writeAttribute('draw:fill-color', '#'.$series->getFill()->getStartColor()->getRGB());
        }
        // > style:graphic-properties
        $this->xmlContent->endElement();
        // style:text-properties
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#'.$series->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $series->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $series->getFont()->getSize().'pt');
        // > style:text-properties
        $this->xmlContent->endElement();
        
        // > style:style
        $this->xmlContent->endElement();
        
        foreach ($series->getDataPointFills() as $idx => $oFill) {
            // style:style
            $this->xmlContent->startElement('style:style');
            $this->xmlContent->writeAttribute('style:name', 'styleSeries'.$this->numSeries.'_'.$idx);
            $this->xmlContent->writeAttribute('style:family', 'chart');
            // style:graphic-properties
            $this->xmlContent->startElement('style:graphic-properties');
            $this->xmlContent->writeAttribute('draw:fill', $oFill->getFillType());
            $this->xmlContent->writeAttribute('draw:fill-color', '#'.$oFill->getStartColor()->getRGB());
            // > style:graphic-properties
            $this->xmlContent->endElement();
            // > style:style
            $this->xmlContent->endElement();
        }
    }
    
    /**
     * @param Chart $chart
     */
    private function writeTable()
    {
        // table:table
        $this->xmlContent->startElement('table:table');
        $this->xmlContent->writeAttribute('table:name', 'table-local');
    
        // table:table-header-columns
        $this->xmlContent->startElement('table:table-header-columns');
        // table:table-column
        $this->xmlContent->startElement('table:table-column');
        // > table:table-column
        $this->xmlContent->endElement();
        // > table:table-header-columns
        $this->xmlContent->endElement();
    
        // table:table-header-rows
        $this->xmlContent->startElement('table:table-header-rows');
        // table:table-row
        $this->xmlContent->startElement('table:table-row');
        if (!empty($this->arrayData)) {
            $rowFirst = reset($this->arrayData);
            foreach ($rowFirst as $key => $cell) {
                // table:table-cell
                $this->xmlContent->startElement('table:table-cell');
                $this->xmlContent->writeAttribute('office:value', 'string');
                // text:p
                $this->xmlContent->startElement('text:p');
                if (isset($this->arrayTitle[$key - 1])) {
                    $this->xmlContent->text($this->arrayTitle[$key - 1]);
                }
                // > text:p
                $this->xmlContent->endElement();
                // > table:table-cell
                $this->xmlContent->endElement();
            }
        }
        // > table:table-row
        $this->xmlContent->endElement();
        // > table:table-header-rows
        $this->xmlContent->endElement();
    
        // table:table-rows
        $this->xmlContent->startElement('table:table-rows');
    
        foreach ($this->arrayData as $row) {
            // table:table-row
            $this->xmlContent->startElement('table:table-row');
            foreach ($row as $cell) {
                // table:table-cell
                $this->xmlContent->startElement('table:table-cell');
                if (is_numeric($cell)) {
                    $this->xmlContent->writeAttribute('office:value-type', 'float');
                    $this->xmlContent->writeAttribute('office:value', $cell);
                } else {
                    $this->xmlContent->writeAttribute('office:value-type', 'string');
                }
                // text:p
                $this->xmlContent->startElement('text:p');
                $this->xmlContent->text($cell);
                // > text:p
                $this->xmlContent->endElement();
                // > table:table-cell
                $this->xmlContent->endElement();
            }
            // > table:table-row
            $this->xmlContent->endElement();
        }
    
        // > table:table-rows
        $this->xmlContent->endElement();
        // > table:table
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     */
    private function writeTitle(Chart $chart)
    {
        // chart:title
        $this->xmlContent->startElement('chart:title');
        $this->xmlContent->writeAttribute('svg:x', String::numberFormat(SharedDrawing::pixelsToCentimeters($chart->getTitle()->getOffsetX()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:y', String::numberFormat(SharedDrawing::pixelsToCentimeters($chart->getTitle()->getOffsetY()), 3) . 'cm');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleTitle');
        // > text:p
        $this->xmlContent->startElement('text:p');
        $this->xmlContent->text($chart->getTitle()->getText());
        // > text:p
        $this->xmlContent->endElement();
        // > chart:title
        $this->xmlContent->endElement();
    }
    
    /**
     * @param Chart $chart
     */
    private function writeTitleStyle(Chart $chart)
    {
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleTitle');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:text-properties
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#'.$chart->getTitle()->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $chart->getTitle()->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $chart->getTitle()->getFont()->getSize().'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $chart->getTitle()->getFont()->isItalic() ? 'italic' : 'normal');
        // > style:text-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }
    
    private function writeWall()
    {
        // chart:wall
        $this->xmlContent->startElement('chart:wall');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleWall');
        // > chart:wall
        $this->xmlContent->endElement();
    }

    /**
     * @param Chart $chart
     */
    private function writeWallStyle(Chart $chart)
    {
        $chartType = $chart->getPlotArea()->getType();
        
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleWall');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:chart-properties
        $this->xmlContent->startElement('style:graphic-properties');
        //@todo : Permit edit color and size border of wall
        if ($chartType instanceof Line || $chartType instanceof Scatter) {
            $this->xmlContent->writeAttribute('draw:fill', 'solid');
            $this->xmlContent->writeAttribute('draw:fill-color', '#FFFFFF');
        } else {
            $this->xmlContent->writeAttribute('draw:fill', 'none');
            $this->xmlContent->writeAttribute('draw:stroke', 'solid');
            $this->xmlContent->writeAttribute('svg:stroke-width', '0.026cm');
            $this->xmlContent->writeAttribute('svg:stroke-color', '#878787');
        }
        // > style:chart-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }
}
