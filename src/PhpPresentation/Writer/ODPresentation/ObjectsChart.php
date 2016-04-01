<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Title;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractTypeBar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractTypePie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;

class ObjectsChart extends AbstractDecoratorWriter
{
    /**
     * @var XMLWriter
     */
    protected $xmlContent;
    /**
     * @var mixed
     */
    protected $arrayData;
    /**
     * @var mixed
     */
    protected $arrayTitle;
    /**
     * @var integer
     */
    protected $numData;
    /**
     * @var integer
     */
    protected $numSeries;
    /**
     * @var string
     */
    protected $rangeCol;

    /**
     * @return ZipInterface
     */
    public function render()
    {
        foreach ($this->getArrayChart() as $keyChart => $shapeChart) {
            $content = $this->writeContentPart($shapeChart);

            if (!empty($content)) {
                $this->getZip()->addFromString('Object '.$keyChart.'/content.xml', $content);
            }
        }

        return $this->getZip();
    }

    /**
     * @param Chart $chart
     * @return string
     * @throws \Exception
     */
    protected function writeContentPart(Chart $chart)
    {
        $this->xmlContent = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        $chartType = $chart->getPlotArea()->getType();

        // Data
        $this->arrayData = array();
        $this->arrayTitle = array();
        $this->numData = 0;
        foreach ($chartType->getSeries() as $series) {
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
        $this->writeAxisStyle($chart);

        // Series
        $this->numSeries = 0;
        foreach ($chartType->getSeries() as $series) {
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
        $this->writeTitleStyle($chart->getTitle());

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
        $this->xmlContent->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters($chart->getWidth()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters($chart->getHeight()), 3) . 'cm');
        $this->xmlContent->writeAttribute('xlink:href', '.');
        $this->xmlContent->writeAttribute('xlink:type', 'simple');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleChart');
        $this->xmlContent->writeAttributeIf($chartType instanceof Area, 'chart:class', 'chart:area');
        $this->xmlContent->writeAttributeIf($chartType instanceof AbstractTypeBar, 'chart:class', 'chart:bar');
        $this->xmlContent->writeAttributeIf($chartType instanceof Line, 'chart:class', 'chart:line');
        $this->xmlContent->writeAttributeIf($chartType instanceof AbstractTypePie, 'chart:class', 'chart:circle');
        $this->xmlContent->writeAttributeIf($chartType instanceof Scatter, 'chart:class', 'chart:scatter');

        //**** Title ****
        $this->writeTitle($chart->getTitle());

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

        return $this->xmlContent->getData();
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
        // chart:axis > chart:categories
        $this->xmlContent->startElement('chart:categories');
        $this->xmlContent->writeAttribute('table:cell-range-address', 'table-local.$A$2:.$A$'.($this->numData+1));
        $this->xmlContent->endElement();
        // chart:axis > chart:grid
        $this->writeGridline($chart->getPlotArea()->getAxisX()->getMajorGridlines(), 'styleAxisXGridlinesMajor', 'major');
        // chart:axis > chart:grid
        $this->writeGridline($chart->getPlotArea()->getAxisX()->getMinorGridlines(), 'styleAxisXGridlinesMinor', 'minor');
        // ##chart:axis
        $this->xmlContent->endElement();

        // chart:axis
        $this->xmlContent->startElement('chart:axis');
        $this->xmlContent->writeAttribute('chart:dimension', 'y');
        $this->xmlContent->writeAttribute('chart:name', 'primary-y');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleAxisY');
        // chart:axis > chart:grid
        $this->writeGridline($chart->getPlotArea()->getAxisY()->getMajorGridlines(), 'styleAxisYGridlinesMajor', 'major');
        // chart:axis > chart:grid
        $this->writeGridline($chart->getPlotArea()->getAxisY()->getMinorGridlines(), 'styleAxisYGridlinesMinor', 'minor');
        // ##chart:axis
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

    protected function writeGridline($oGridlines, $styleName, $chartClass)
    {
        if (!($oGridlines instanceof Chart\Gridlines)) {
            return ;
        }

        $this->xmlContent->startElement('chart:grid');
        $this->xmlContent->writeAttribute('chart:style-name', $styleName);
        $this->xmlContent->writeAttribute('chart:class', $chartClass);
        $this->xmlContent->endElement();
    }

    /**
     * @param Chart $chart
     * @todo Set function in \PhpPresentation\Shape\Chart\Axis for defining width and color of the axis
     */
    protected function writeAxisStyle(Chart $chart)
    {
        $chartType = $chart->getPlotArea()->getType();

        // AxisX
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleAxisX');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:style > style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        $this->xmlContent->writeAttribute('chart:display-label', 'true');
        $this->xmlContent->writeAttribute('chart:tick-marks-major-inner', 'false');
        $this->xmlContent->writeAttribute('chart:tick-marks-major-outer', 'false');
        if ($chartType instanceof AbstractTypePie) {
            $this->xmlContent->writeAttribute('chart:reverse-direction', 'true');
        }
        $this->xmlContent->endElement();
        // style:style > style:text-properties
        $oFont = $chart->getPlotArea()->getAxisX()->getFont();
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#'.$oFont->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $oFont->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $oFont->getSize().'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $oFont->isItalic() ? 'italic' : 'normal');
        $this->xmlContent->endElement();
        // style:style > style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        $this->xmlContent->writeAttribute('svg:stroke-width', '0.026cm');
        $this->xmlContent->writeAttribute('svg:stroke-color', '#878787');
        $this->xmlContent->endElement();
        // ##style:style
        $this->xmlContent->endElement();

        // AxisX GridLines Major
        $this->writeGridlineStyle($chart->getPlotArea()->getAxisX()->getMajorGridlines(), 'styleAxisXGridlinesMajor');

        // AxisX GridLines Minor
        $this->writeGridlineStyle($chart->getPlotArea()->getAxisX()->getMinorGridlines(), 'styleAxisXGridlinesMinor');

        // AxisY
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleAxisY');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:style > style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        $this->xmlContent->writeAttribute('chart:display-label', 'true');
        $this->xmlContent->writeAttribute('chart:tick-marks-major-inner', 'false');
        $this->xmlContent->writeAttribute('chart:tick-marks-major-outer', 'false');
        if ($chartType instanceof AbstractTypePie) {
            $this->xmlContent->writeAttribute('chart:reverse-direction', 'true');
        }
        $this->xmlContent->endElement();
        // style:style > style:text-properties
        $oFont = $chart->getPlotArea()->getAxisY()->getFont();
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#'.$oFont->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $oFont->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $oFont->getSize().'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $oFont->isItalic() ? 'italic' : 'normal');
        $this->xmlContent->endElement();
        // style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        $this->xmlContent->writeAttribute('svg:stroke-width', '0.026cm');
        $this->xmlContent->writeAttribute('svg:stroke-color', '#878787');
        $this->xmlContent->endElement();
        // ## style:style
        $this->xmlContent->endElement();

        // AxisY GridLines Major
        $this->writeGridlineStyle($chart->getPlotArea()->getAxisY()->getMajorGridlines(), 'styleAxisYGridlinesMajor');

        // AxisY GridLines Minor
        $this->writeGridlineStyle($chart->getPlotArea()->getAxisY()->getMinorGridlines(), 'styleAxisYGridlinesMinor');
    }

    /**
     * @param Chart\Gridlines $oGridlines
     * @param string $styleName
     */
    protected function writeGridlineStyle($oGridlines, $styleName)
    {
        if (!($oGridlines instanceof Chart\Gridlines)) {
            return;
        }
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', $styleName);
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:style > style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        $this->xmlContent->writeAttribute('svg:stroke-width', number_format(CommonDrawing::pointsToCentimeters($oGridlines->getOutline()->getWidth()), 2, '.', '').'cm');
        $this->xmlContent->writeAttribute('svg:stroke-color', '#'.$oGridlines->getOutline()->getFill()->getStartColor()->getRGB());
        $this->xmlContent->endElement();
        // ##style:style
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
        $this->xmlContent->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters($chart->getLegend()->getOffsetX()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters($chart->getLegend()->getOffsetY()), 3) . 'cm');
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
        foreach ($chartType->getSeries() as $series) {
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
     * @link : http://books.evc-cit.info/odbook/ch08.html#chart-plot-area-section
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
        if ($chartType instanceof Bar3D) {
            $this->xmlContent->writeAttribute('chart:three-dimensional', 'true');
            $this->xmlContent->writeAttribute('chart:right-angled-axes', 'true');
        } elseif ($chartType instanceof Pie3D) {
            $this->xmlContent->writeAttribute('chart:three-dimensional', 'true');
            $this->xmlContent->writeAttribute('chart:right-angled-axes', 'true');
        }
        if ($chartType instanceof AbstractTypeBar) {
            $chartVertical = 'false';
            if ($chartType->getBarDirection() == AbstractTypeBar::DIRECTION_HORIZONTAL) {
                $chartVertical = 'true';
            }
            $this->xmlContent->writeAttribute('chart:vertical', $chartVertical);
            if ($chartType->getBarGrouping() == Bar::GROUPING_CLUSTERED) {
                $this->xmlContent->writeAttribute('chart:stacked', 'false');
                $this->xmlContent->writeAttribute('chart:overlap', '0');
            } elseif ($chartType->getBarGrouping() == Bar::GROUPING_STACKED) {
                $this->xmlContent->writeAttribute('chart:stacked', 'true');
                $this->xmlContent->writeAttribute('chart:overlap', '100');
            } elseif ($chartType->getBarGrouping() == Bar::GROUPING_PERCENTSTACKED) {
                $this->xmlContent->writeAttribute('chart:stacked', 'true');
                $this->xmlContent->writeAttribute('chart:overlap', '100');
                $this->xmlContent->writeAttribute('chart:percentage', 'true');
            }
        }
        $labelFormat = 'value';
        if ($chartType instanceof AbstractTypeBar) {
            if ($chartType->getBarGrouping() == Bar::GROUPING_PERCENTSTACKED) {
                $labelFormat = 'percentage';
            }
        }
        $this->xmlContent->writeAttribute('chart:data-label-number', $labelFormat);

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
        if ($chartType instanceof Area) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:area');
        } elseif ($chartType instanceof AbstractTypeBar) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:bar');
        } elseif ($chartType instanceof Line) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:line');
        } elseif ($chartType instanceof AbstractTypePie) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:circle');
        } elseif ($chartType instanceof Scatter) {
            $this->xmlContent->writeAttribute('chart:class', 'chart:scatter');
        }
        $this->xmlContent->writeAttribute('chart:style-name', 'styleSeries'.$this->numSeries);
        if ($chartType instanceof Area || $chartType instanceof AbstractTypeBar || $chartType instanceof Line || $chartType instanceof Scatter) {
            $dataPointFills = $series->getDataPointFills();

            $incRepeat = $numRange;
            if (!empty($dataPointFills)) {
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
        } elseif ($chartType instanceof AbstractTypePie) {
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
        $this->xmlContent->writeAttribute('chart:label-position', 'center');
        if ($chartType instanceof AbstractTypePie) {
            $this->xmlContent->writeAttribute('chart:pie-offset', $chartType->getExplosion());
        }
        if ($chartType instanceof Line || $chartType instanceof Scatter) {
            $oMarker = $series->getMarker();
            /**
             * @link : http://www.datypic.com/sc/odf/a-chart_symbol-type.html
             */
            $this->xmlContent->writeAttributeIf($oMarker->getSymbol() == Chart\Marker::SYMBOL_NONE, 'chart:symbol-type', 'none');
            /**
             * @link : http://www.datypic.com/sc/odf/a-chart_symbol-name.html
             */
            $this->xmlContent->writeAttributeIf($oMarker->getSymbol() != Chart\Marker::SYMBOL_NONE, 'chart:symbol-type', 'named-symbol');
            if ($oMarker->getSymbol() != Chart\Marker::SYMBOL_NONE) {
                switch ($oMarker->getSymbol()) {
                    case Chart\Marker::SYMBOL_DASH:
                        $symbolName = 'horizontal-bar';
                        break;
                    case Chart\Marker::SYMBOL_DOT:
                        $symbolName = 'circle';
                        break;
                    case Chart\Marker::SYMBOL_TRIANGLE:
                        $symbolName = 'arrow-up';
                        break;
                    default:
                        $symbolName = $oMarker->getSymbol();
                        break;
                }
                $this->xmlContent->writeAttribute('chart:symbol-name', $symbolName);
                $symbolSize = number_format(CommonDrawing::pointsToCentimeters($oMarker->getSize()), 2, '.', '');
                $this->xmlContent->writeAttribute('chart:symbol-width', $symbolSize.'cm');
                $this->xmlContent->writeAttribute('chart:symbol-height', $symbolSize.'cm');
            }
        }
        // > style:chart-properties
        $this->xmlContent->endElement();
        // style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        if ($chartType instanceof Line || $chartType instanceof Scatter) {
            $outlineWidth = '';
            $outlineColor = '';

            $oOutline = $series->getOutline();
            if ($oOutline instanceof Outline) {
                $outlineWidth = $oOutline->getWidth();
                if (!empty($outlineWidth)) {
                    $outlineWidth = number_format(CommonDrawing::pointsToCentimeters($outlineWidth), 3, '.', '');
                }
                $outlineColor = $oOutline->getFill()->getStartColor()->getRGB();
            }
            if (empty($outlineWidth)) {
                $outlineWidth = '0.079';
            }
            if (empty($outlineColor)) {
                $outlineColor = '4a7ebb';
            }
            $this->xmlContent->writeAttribute('svg:stroke-width', $outlineWidth.'cm');
            $this->xmlContent->writeAttribute('svg:stroke-color', '#'.$outlineColor);
        } else {
            $this->xmlContent->writeAttribute('draw:stroke', 'none');
            if (!($chartType instanceof Area)) {
                $this->xmlContent->writeAttribute('draw:fill', $series->getFill()->getFillType());
            }
        }
        $this->xmlContent->writeAttribute('draw:fill-color', '#'.$series->getFill()->getStartColor()->getRGB());
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

        // table:table-columns
        $this->xmlContent->startElement('table:table-columns');
        // table:table-column
        $this->xmlContent->startElement('table:table-column');
        if (!empty($this->arrayData)) {
            $rowFirst = reset($this->arrayData);
            $this->xmlContent->writeAttribute('table:number-columns-repeated', count($rowFirst) - 1);
        }
        // > table:table-column
        $this->xmlContent->endElement();
        // > table:table-columns
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
                if (isset($this->arrayTitle[$key - 1])) {
                    $this->xmlContent->writeAttribute('office:value-type', 'string');
                }
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

                $cellNumeric = is_numeric($cell);
                $this->xmlContent->writeAttributeIf(!$cellNumeric, 'office:value-type', 'string');
                $this->xmlContent->writeAttributeIf($cellNumeric, 'office:value-type', 'float');
                $this->xmlContent->writeAttributeIf($cellNumeric, 'office:value', $cell);
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
     * @param Title $oTitle
     */
    private function writeTitle(Title $oTitle)
    {
        if (!$oTitle->isVisible()) {
            return;
        }
        // chart:title
        $this->xmlContent->startElement('chart:title');
        $this->xmlContent->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters($oTitle->getOffsetX()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters($oTitle->getOffsetY()), 3) . 'cm');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleTitle');
        // > text:p
        $this->xmlContent->startElement('text:p');
        $this->xmlContent->text($oTitle->getText());
        // > text:p
        $this->xmlContent->endElement();
        // > chart:title
        $this->xmlContent->endElement();
    }

    /**
     * @param Title $oTitle
     */
    private function writeTitleStyle(Title $oTitle)
    {
        if (!$oTitle->isVisible()) {
            return;
        }
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleTitle');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:text-properties
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#'.$oTitle->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $oTitle->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $oTitle->getFont()->getSize().'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $oTitle->getFont()->isItalic() ? 'italic' : 'normal');
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
