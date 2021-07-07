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

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\Title;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractTypeBar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractTypeLine;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractTypePie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Doughnut;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Radar;
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
     * @var int
     */
    protected $numData;

    /**
     * @var int
     */
    protected $numSeries;

    /**
     * @var string
     */
    protected $rangeCol;

    public function render(): ZipInterface
    {
        foreach ($this->getArrayChart() as $keyChart => $shapeChart) {
            $content = $this->writeContentPart($shapeChart);

            if (!empty($content)) {
                $this->getZip()->addFromString('Object ' . $keyChart . '/content.xml', $content);
            }
        }

        return $this->getZip();
    }

    protected function writeContentPart(Chart $chart): string
    {
        $this->xmlContent = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        $chartType = $chart->getPlotArea()->getType();

        // Data
        $this->arrayData = [];
        $this->arrayTitle = [];
        $this->numData = 0;
        foreach ($chartType->getSeries() as $series) {
            $inc = 0;
            $this->arrayTitle[] = $series->getTitle();
            foreach ($series->getValues() as $key => $value) {
                if (!isset($this->arrayData[$inc])) {
                    $this->arrayData[$inc] = [];
                }
                if (empty($this->arrayData[$inc])) {
                    $this->arrayData[$inc][] = $key;
                }
                $this->arrayData[$inc][] = $value;
                ++$inc;
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

        // Styles
        $this->writeChartStyle($chart);
        $this->writeAxisStyle($chart);
        $this->numSeries = 0;
        foreach ($chartType->getSeries() as $series) {
            $this->writeSeriesStyle($chart, $series);
            ++$this->numSeries;
        }
        $this->writeFloorStyle();
        $this->writeLegendStyle($chart);
        $this->writePlotAreaStyle($chart);
        $this->writeTitleStyle($chart->getTitle());
        $this->writeWallStyle($chart);

        // > office:automatic-styles
        $this->xmlContent->endElement();

        // office:body
        $this->xmlContent->startElement('office:body');
        // office:chart
        $this->xmlContent->startElement('office:chart');
        // office:chart
        $this->xmlContent->startElement('chart:chart');
        $this->xmlContent->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $chart->getWidth()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $chart->getHeight()), 3) . 'cm');
        $this->xmlContent->writeAttribute('xlink:href', '.');
        $this->xmlContent->writeAttribute('xlink:type', 'simple');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleChart');
        $this->xmlContent->writeAttributeIf($chartType instanceof Area, 'chart:class', 'chart:area');
        $this->xmlContent->writeAttributeIf($chartType instanceof AbstractTypeBar, 'chart:class', 'chart:bar');
        if (!($chartType instanceof Doughnut)) {
            $this->xmlContent->writeAttributeIf($chartType instanceof AbstractTypePie, 'chart:class', 'chart:circle');
        }
        $this->xmlContent->writeAttributeIf($chartType instanceof Doughnut, 'chart:class', 'chart:ring');
        $this->xmlContent->writeAttributeIf($chartType instanceof Line, 'chart:class', 'chart:line');
        $this->xmlContent->writeAttributeIf($chartType instanceof Radar, 'chart:class', 'chart:radar');
        $this->xmlContent->writeAttributeIf($chartType instanceof Scatter, 'chart:class', 'chart:scatter');

        $this->writeTitle($chart->getTitle());
        $this->writeLegend($chart);
        $this->writePlotArea($chart);
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

    protected function writeAxis(Chart $chart): void
    {
        $chartType = $chart->getPlotArea()->getType();

        // chart:axis
        $this->xmlContent->startElement('chart:axis');
        $this->xmlContent->writeAttribute('chart:dimension', 'x');
        $this->xmlContent->writeAttribute('chart:name', 'primary-x');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleAxisX');
        // chart:axis > chart:title
        if ($chart->getPlotArea()->getAxisX()->isVisible()) {
            $this->xmlContent->startElement('chart:title');
            $this->xmlContent->writeAttribute('chart:style-name', 'styleAxisXTitle');
            $this->xmlContent->writeElement('text:p', $chart->getPlotArea()->getAxisX()->getTitle());
            $this->xmlContent->endElement();
        }
        // chart:axis > chart:categories
        $this->xmlContent->startElement('chart:categories');
        $this->xmlContent->writeAttribute('table:cell-range-address', 'table-local.$A$2:.$A$' . ($this->numData + 1));
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
        // chart:axis > chart:title
        if ($chart->getPlotArea()->getAxisY()->isVisible()) {
            $this->xmlContent->startElement('chart:title');
            $this->xmlContent->writeAttribute('chart:style-name', 'styleAxisYTitle');
            $this->xmlContent->writeElement('text:p', $chart->getPlotArea()->getAxisY()->getTitle());
            $this->xmlContent->endElement();
        }
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

    protected function writeGridline(?Chart\Gridlines $oGridlines, string $styleName, string $chartClass): void
    {
        if (!$oGridlines) {
            return;
        }

        $this->xmlContent->startElement('chart:grid');
        $this->xmlContent->writeAttribute('chart:style-name', $styleName);
        $this->xmlContent->writeAttribute('chart:class', $chartClass);
        $this->xmlContent->endElement();
    }

    /**
     * @todo Set function in \PhpPresentation\Shape\Chart\Axis for defining width and color of the axis
     */
    protected function writeAxisStyle(Chart $chart): void
    {
        $chartType = $chart->getPlotArea()->getType();

        // AxisX
        $this->writeAxisMainStyle($chart->getPlotArea()->getAxisX(), 'styleAxisX', $chartType);

        // AxisX Title
        $this->writeAxisTitleStyle($chart->getPlotArea()->getAxisX(), 'styleAxisXTitle');

        // AxisX GridLines Major
        $this->writeGridlineStyle($chart->getPlotArea()->getAxisX()->getMajorGridlines(), 'styleAxisXGridlinesMajor');

        // AxisX GridLines Minor
        $this->writeGridlineStyle($chart->getPlotArea()->getAxisX()->getMinorGridlines(), 'styleAxisXGridlinesMinor');

        // AxisY
        $this->writeAxisMainStyle($chart->getPlotArea()->getAxisY(), 'styleAxisY', $chartType);

        // AxisY Title
        $this->writeAxisTitleStyle($chart->getPlotArea()->getAxisY(), 'styleAxisYTitle');

        // AxisY GridLines Major
        $this->writeGridlineStyle($chart->getPlotArea()->getAxisY()->getMajorGridlines(), 'styleAxisYGridlinesMajor');

        // AxisY GridLines Minor
        $this->writeGridlineStyle($chart->getPlotArea()->getAxisY()->getMinorGridlines(), 'styleAxisYGridlinesMinor');
    }

    protected function writeAxisMainStyle(Axis $axis, string $styleName, AbstractType $chartType): void
    {
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', $styleName);
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:style > style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        $this->xmlContent->writeAttribute('chart:display-label', 'true');
        $this->xmlContent->writeAttribute('chart:tick-marks-major-inner', 'false');
        $this->xmlContent->writeAttribute('chart:tick-marks-major-outer', 'false');
        $this->xmlContent->writeAttributeIf($chartType instanceof AbstractTypePie, 'chart:reverse-direction', 'true');
        $this->xmlContent->writeAttributeIf(null !== $axis->getMinBounds(), 'chart:minimum', $axis->getMinBounds());
        $this->xmlContent->writeAttributeIf(null !== $axis->getMaxBounds(), 'chart:maximum', $axis->getMaxBounds());
        $this->xmlContent->writeAttributeIf(null !== $axis->getMajorUnit(), 'chart:interval-major', $axis->getMajorUnit());
        $this->xmlContent->writeAttributeIf(null !== $axis->getMinorUnit(), 'chart:interval-minor-divisor', $axis->getMinorUnit());
        switch ($axis->getTickLabelPosition()) {
            case Axis::TICK_LABEL_POSITION_NEXT_TO:
                $this->xmlContent->writeAttribute('chart:axis-label-position', 'near-axis');

                break;
            case Axis::TICK_LABEL_POSITION_HIGH:
                $this->xmlContent->writeAttribute('chart:axis-position', '0');
                $this->xmlContent->writeAttribute('chart:axis-label-position', 'outside-end');

                break;
            case Axis::TICK_LABEL_POSITION_LOW:
                $this->xmlContent->writeAttribute('chart:axis-position', '0');
                $this->xmlContent->writeAttribute('chart:axis-label-position', 'outside-start');
                $this->xmlContent->writeAttribute('chart:tick-mark-position', 'at-axis');

                break;
        }
        $this->xmlContent->writeAttributeIf($chartType instanceof Radar && $styleName == 'styleAxisX', 'chart:reverse-direction', 'true');
        $this->xmlContent->endElement();
        // style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        if ($axis->getOutline()->getFill()->getFillType() === Fill::FILL_NONE) {
            $this->xmlContent->writeAttribute('draw:stroke', 'none');
        } else {
            $this->xmlContent->writeAttribute('draw:stroke', 'solid');
        }
        $this->xmlContent->writeAttribute('svg:stroke-width', number_format(CommonDrawing::pointsToCentimeters($axis->getOutline()->getWidth()), 3, '.', '') . 'cm');
        $this->xmlContent->writeAttribute('svg:stroke-color', '#' . $axis->getOutline()->getFill()->getStartColor()->getRGB());
        $this->xmlContent->endElement();
        // style:style > style:text-properties
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#' . $axis->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $axis->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $axis->getFont()->getSize() . 'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $axis->getFont()->isItalic() ? 'italic' : 'normal');
        $this->xmlContent->endElement();
        // ## style:style
        $this->xmlContent->endElement();
    }

    protected function writeAxisTitleStyle(Axis $axis, string $styleName): void
    {
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', $styleName);
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        $this->xmlContent->writeAttribute('chart:auto-position', 'true');
        $this->xmlContent->writeAttributeIf($axis->getTitleRotation() != 0, 'style:rotation-angle', '-' . $axis->getTitleRotation());
        // > style:chart-properties
        $this->xmlContent->endElement();
        // style:text-properties
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#' . $axis->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $axis->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $axis->getFont()->getSize() . 'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $axis->getFont()->isItalic() ? 'italic' : 'normal');
        // > style:text-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }

    protected function writeGridlineStyle(?Chart\Gridlines $oGridlines, string $styleName): void
    {
        if (!$oGridlines) {
            return;
        }
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', $styleName);
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:style > style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        $this->xmlContent->writeAttribute('svg:stroke-width', number_format(CommonDrawing::pointsToCentimeters($oGridlines->getOutline()->getWidth()), 2, '.', '') . 'cm');
        $this->xmlContent->writeAttribute('svg:stroke-color', '#' . $oGridlines->getOutline()->getFill()->getStartColor()->getRGB());
        $this->xmlContent->endElement();
        // ##style:style
        $this->xmlContent->endElement();
    }

    protected function writeChartStyle(Chart $chart): void
    {
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleChart');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        if ($chart->getFill()->getFillType() === Fill::FILL_NONE) {
            $this->xmlContent->writeAttribute('draw:stroke', 'none');
        } else {
            $this->xmlContent->writeAttribute('draw:stroke', 'solid');
        }
        $this->xmlContent->writeAttribute('draw:fill-color', '#' . $chart->getFill()->getStartColor()->getRGB());
        // > style:graphic-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }

    protected function writeFloor(): void
    {
        // chart:floor
        $this->xmlContent->startElement('chart:floor');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleFloor');
        // > chart:floor
        $this->xmlContent->endElement();
    }

    protected function writeFloorStyle(): void
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

    protected function writeLegend(Chart $chart): void
    {
        // chart:legend
        $this->xmlContent->startElement('chart:legend');
        switch ($chart->getLegend()->getPosition()) {
            case Chart\Legend::POSITION_BOTTOM:
                $position = 'bottom';

                break;
            case Chart\Legend::POSITION_LEFT:
                $position = 'start';

                break;
            case Chart\Legend::POSITION_TOP:
                $position = 'top';

                break;
            case Chart\Legend::POSITION_TOPRIGHT:
                $position = 'top-end';

                break;
            case Chart\Legend::POSITION_RIGHT:
            default:
                $position = 'end';

                break;
        }
        $this->xmlContent->writeAttribute('chart:legend-position', $position);
        $this->xmlContent->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $chart->getLegend()->getOffsetX()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $chart->getLegend()->getOffsetY()), 3) . 'cm');
        $this->xmlContent->writeAttribute('style:legend-expansion', 'high');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleLegend');
        // > chart:legend
        $this->xmlContent->endElement();
    }

    protected function writeLegendStyle(Chart $chart): void
    {
        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleLegend');
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        $this->xmlContent->writeAttribute('chart:auto-position', 'true');
        // > style:chart-properties
        $this->xmlContent->endElement();
        // style:text-properties
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#' . $chart->getLegend()->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $chart->getLegend()->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $chart->getLegend()->getFont()->getSize() . 'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $chart->getLegend()->getFont()->isItalic() ? 'italic' : 'normal');
        // > style:text-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }

    protected function writePlotArea(Chart $chart): void
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
            $arrayLight = [
                ['#808080', '(0 0 1)', 'false', 'true'],
                ['#666666', '(0.2 0.4 1)', 'true', 'false'],
                ['#808080', '(0 0 1)', 'false', 'false'],
                ['#808080', '(0 0 1)', 'false', 'false'],
                ['#808080', '(0 0 1)', 'false', 'false'],
                ['#808080', '(0 0 1)', 'false', 'false'],
                ['#808080', '(0 0 1)', 'false', 'false'],
            ];
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
            ++$this->rangeCol;
            ++$this->numSeries;
        }

        //**** Wall ****
        $this->writeWall();
        //**** Floor ****
        $this->writeFloor();
        // > chart:plot-area
        $this->xmlContent->endElement();
    }

    /**
     * @see : http://books.evc-cit.info/odbook/ch08.html#chart-plot-area-section
     */
    protected function writePlotAreaStyle(Chart $chart): void
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
        } elseif ($chartType instanceof AbstractTypeLine) {
            $this->xmlContent->writeAttributeIf($chartType->isSmooth(), 'chart:interpolation', 'cubic-spline');
        }
        switch ($chart->getDisplayBlankAs()) {
            case Chart::BLANKAS_ZERO:
                $this->xmlContent->writeAttribute('chart:treat-empty-cells', 'use-zero');

                break;
            case Chart::BLANKAS_GAP:
                $this->xmlContent->writeAttribute('chart:treat-empty-cells', 'leave-gap');

                break;
            case Chart::BLANKAS_SPAN:
                $this->xmlContent->writeAttribute('chart:treat-empty-cells', 'ignore');

                break;
        }
        if ($chartType instanceof AbstractTypeBar) {
            $chartVertical = 'false';
            if (AbstractTypeBar::DIRECTION_HORIZONTAL == $chartType->getBarDirection()) {
                $chartVertical = 'true';
            }
            $this->xmlContent->writeAttribute('chart:vertical', $chartVertical);
            if (Bar::GROUPING_CLUSTERED == $chartType->getBarGrouping()) {
                $this->xmlContent->writeAttribute('chart:stacked', 'false');
                $this->xmlContent->writeAttribute('chart:overlap', '0');
            } elseif (Bar::GROUPING_STACKED == $chartType->getBarGrouping()) {
                $this->xmlContent->writeAttribute('chart:stacked', 'true');
                $this->xmlContent->writeAttribute('chart:overlap', '100');
            } elseif (Bar::GROUPING_PERCENTSTACKED == $chartType->getBarGrouping()) {
                $this->xmlContent->writeAttribute('chart:stacked', 'true');
                $this->xmlContent->writeAttribute('chart:overlap', '100');
                $this->xmlContent->writeAttribute('chart:percentage', 'true');
            }
        }
        $labelFormat = 'value';
        if ($chartType instanceof AbstractTypeBar) {
            if (Bar::GROUPING_PERCENTSTACKED == $chartType->getBarGrouping()) {
                $labelFormat = 'percentage';
            }
        }
        $this->xmlContent->writeAttribute('chart:data-label-number', $labelFormat);

        // > style:text-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }

    protected function writeSeries(Chart $chart, Chart\Series $series): void
    {
        $chartType = $chart->getPlotArea()->getType();

        $numRange = count($series->getValues());
        // chart:series
        $this->xmlContent->startElement('chart:series');
        $this->xmlContent->writeAttribute('chart:values-cell-range-address', 'table-local.$' . $this->rangeCol . '$2:.$' . $this->rangeCol . '$' . ($numRange + 1));
        $this->xmlContent->writeAttribute('chart:label-cell-address', 'table-local.$' . $this->rangeCol . '$1');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleSeries' . $this->numSeries);
        if ($chartType instanceof Area
            || $chartType instanceof AbstractTypeBar
            || $chartType instanceof Line
            || $chartType instanceof Radar
            || $chartType instanceof Scatter
        ) {
            $dataPointFills = $series->getDataPointFills();

            $incRepeat = $numRange;
            if (!empty($dataPointFills)) {
                $inc = 0;
                $incRepeat = 1;
                $newFill = new Fill();
                do {
                    if ($series->getDataPointFill($inc)->getHashCode() !== $newFill->getHashCode()) {
                        // chart:data-point
                        $this->xmlContent->startElement('chart:data-point');
                        $this->xmlContent->writeAttribute('chart:repeated', $incRepeat);
                        // > chart:data-point
                        $this->xmlContent->endElement();
                        $incRepeat = 1;

                        // chart:data-point
                        $this->xmlContent->startElement('chart:data-point');
                        $this->xmlContent->writeAttribute('chart:style-name', 'styleSeries' . $this->numSeries . '_' . $inc);
                        // > chart:data-point
                        $this->xmlContent->endElement();
                    }
                    ++$inc;
                    ++$incRepeat;
                } while ($inc < $numRange);
                --$incRepeat;
            }
            // chart:data-point
            $this->xmlContent->startElement('chart:data-point');
            $this->xmlContent->writeAttribute('chart:repeated', $incRepeat);
            // > chart:data-point
            $this->xmlContent->endElement();
        } elseif ($chartType instanceof AbstractTypePie) {
            $count = count($series->getDataPointFills());
            for ($inc = 0; $inc < $count; ++$inc) {
                // chart:data-point
                $this->xmlContent->startElement('chart:data-point');
                $this->xmlContent->writeAttribute('chart:style-name', 'styleSeries' . $this->numSeries . '_' . $inc);
                // > chart:data-point
                $this->xmlContent->endElement();
            }
        }

        // > chart:series
        $this->xmlContent->endElement();
    }

    protected function writeSeriesStyle(Chart $chart, Chart\Series $series): void
    {
        $chartType = $chart->getPlotArea()->getType();

        // style:style
        $this->xmlContent->startElement('style:style');
        $this->xmlContent->writeAttribute('style:name', 'styleSeries' . $this->numSeries);
        $this->xmlContent->writeAttribute('style:family', 'chart');
        // style:chart-properties
        $this->xmlContent->startElement('style:chart-properties');
        if ($series->hasShowValue()) {
            if ($series->hasShowPercentage()) {
                $this->xmlContent->writeAttribute('chart:data-label-number', 'value-and-percentage');
            } else {
                $this->xmlContent->writeAttribute('chart:data-label-number', 'value');
            }
        } else {
            if ($series->hasShowPercentage()) {
                $this->xmlContent->writeAttribute('chart:data-label-number', 'percentage');
            }
        }
        if ($series->hasShowCategoryName()) {
            $this->xmlContent->writeAttribute('chart:data-label-text', 'true');
        }
        $this->xmlContent->writeAttribute('chart:label-position', 'center');
        if ($chartType instanceof AbstractTypePie) {
            $this->xmlContent->writeAttribute('chart:pie-offset', $chartType->getExplosion());
        }
        if ($chartType instanceof Line || $chartType instanceof Scatter) {
            $oMarker = $series->getMarker();
            // @link : http://www.datypic.com/sc/odf/a-chart_symbol-type.html
            $this->xmlContent->writeAttributeIf(Chart\Marker::SYMBOL_NONE == $oMarker->getSymbol(), 'chart:symbol-type', 'none');
            // @link : http://www.datypic.com/sc/odf/a-chart_symbol-name.html
            $this->xmlContent->writeAttributeIf(Chart\Marker::SYMBOL_NONE != $oMarker->getSymbol(), 'chart:symbol-type', 'named-symbol');
            if (Chart\Marker::SYMBOL_NONE != $oMarker->getSymbol()) {
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
                $this->xmlContent->writeAttribute('chart:symbol-width', $symbolSize . 'cm');
                $this->xmlContent->writeAttribute('chart:symbol-height', $symbolSize . 'cm');
            }
        }

        $separator = $series->getSeparator();
        if (!empty($separator)) {
            // style:chart-properties/chart:label-separator
            $this->xmlContent->startElement('chart:label-separator');
            if (PHP_EOL == $separator) {
                $this->xmlContent->writeRaw('<text:p><text:line-break /></text:p>');
            } else {
                $this->xmlContent->writeElement('text:p', $separator);
            }
            $this->xmlContent->endElement();
        }

        // > style:chart-properties
        $this->xmlContent->endElement();
        // style:graphic-properties
        $this->xmlContent->startElement('style:graphic-properties');
        if ($chartType instanceof Line || $chartType instanceof Radar || $chartType instanceof Scatter) {
            $outlineWidth = '';
            $outlineColor = '';

            $oOutline = $series->getOutline();
            if ($oOutline instanceof Outline) {
                $outlineWidth = $oOutline->getWidth();
                if (!empty($outlineWidth)) {
                    $outlineWidth = number_format(CommonDrawing::pixelsToCentimeters($outlineWidth), 3, '.', '');
                }
                $outlineColor = $oOutline->getFill()->getStartColor()->getRGB();
            }
            if (empty($outlineWidth)) {
                $outlineWidth = '0.079';
            }
            if (empty($outlineColor)) {
                $outlineColor = '4a7ebb';
            }
            $this->xmlContent->writeAttribute('svg:stroke-width', $outlineWidth . 'cm');
            $this->xmlContent->writeAttribute('svg:stroke-color', '#' . $outlineColor);
        } else {
            $this->xmlContent->writeAttribute('draw:stroke', 'none');
            if (!($chartType instanceof Area)) {
                $this->xmlContent->writeAttribute('draw:fill', $series->getFill()->getFillType());
            }
        }
        $this->xmlContent->writeAttribute('draw:fill-color', '#' . $series->getFill()->getStartColor()->getRGB());
        // > style:graphic-properties
        $this->xmlContent->endElement();
        // style:text-properties
        $this->xmlContent->startElement('style:text-properties');
        $this->xmlContent->writeAttribute('fo:color', '#' . $series->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $series->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $series->getFont()->getSize() . 'pt');
        // > style:text-properties
        $this->xmlContent->endElement();

        // > style:style
        $this->xmlContent->endElement();

        foreach ($series->getDataPointFills() as $idx => $oFill) {
            // style:style
            $this->xmlContent->startElement('style:style');
            $this->xmlContent->writeAttribute('style:name', 'styleSeries' . $this->numSeries . '_' . $idx);
            $this->xmlContent->writeAttribute('style:family', 'chart');
            // style:graphic-properties
            $this->xmlContent->startElement('style:graphic-properties');
            $this->xmlContent->writeAttribute('draw:fill', $oFill->getFillType());
            $this->xmlContent->writeAttribute('draw:fill-color', '#' . $oFill->getStartColor()->getRGB());
            // > style:graphic-properties
            $this->xmlContent->endElement();
            // > style:style
            $this->xmlContent->endElement();
        }
    }

    protected function writeTable(): void
    {
        // table:table
        $this->xmlContent->startElement('table:table');
        $this->xmlContent->writeAttribute('table:name', 'table-local');

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

        // table:table-header-columns
        $this->xmlContent->startElement('table:table-header-columns');
        // table:table-column
        $this->xmlContent->writeElement('table:table-column');
        // > table:table-header-columns
        $this->xmlContent->endElement();

        // table:table-header-rows
        $this->xmlContent->startElement('table:table-header-rows');
        // table:table-row
        $this->xmlContent->startElement('table:table-row');
        if (empty($this->arrayData)) {
            $this->xmlContent->startElement('table:table-cell');
            $this->xmlContent->endElement();
        } else {
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
        if (empty($this->arrayData)) {
            $this->xmlContent->startElement('table:table-row');
            $this->xmlContent->startElement('table:table-cell');
            $this->xmlContent->endElement();
            $this->xmlContent->endElement();
        } else {
            foreach ($this->arrayData as $row) {
                // table:table-row
                $this->xmlContent->startElement('table:table-row');
                foreach ($row as $cell) {
                    // table:table-cell
                    $this->xmlContent->startElement('table:table-cell');

                    $cellValueTypeFloat = null === $cell ? true : is_numeric($cell);
                    $this->xmlContent->writeAttributeIf(!$cellValueTypeFloat, 'office:value-type', 'string');
                    $this->xmlContent->writeAttributeIf($cellValueTypeFloat, 'office:value-type', 'float');
                    $this->xmlContent->writeAttributeIf($cellValueTypeFloat, 'office:value', null === $cell ? 'NaN' : $cell);
                    // text:p
                    $this->xmlContent->startElement('text:p');
                    $this->xmlContent->text(null === $cell ? 'NaN' : (string) $cell);
                    $this->xmlContent->endElement();
                    // > table:table-cell
                    $this->xmlContent->endElement();
                }
                // > table:table-row
                $this->xmlContent->endElement();
            }
        }
        // > table:table-rows
        $this->xmlContent->endElement();

        // > table:table
        $this->xmlContent->endElement();
    }

    protected function writeTitle(Title $oTitle): void
    {
        if (!$oTitle->isVisible()) {
            return;
        }
        // chart:title
        $this->xmlContent->startElement('chart:title');
        $this->xmlContent->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $oTitle->getOffsetX()), 3) . 'cm');
        $this->xmlContent->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $oTitle->getOffsetY()), 3) . 'cm');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleTitle');
        // > text:p
        $this->xmlContent->startElement('text:p');
        $this->xmlContent->text($oTitle->getText());
        $this->xmlContent->endElement();
        // > chart:title
        $this->xmlContent->endElement();
    }

    protected function writeTitleStyle(Title $oTitle): void
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
        $this->xmlContent->writeAttribute('fo:color', '#' . $oTitle->getFont()->getColor()->getRGB());
        $this->xmlContent->writeAttribute('fo:font-family', $oTitle->getFont()->getName());
        $this->xmlContent->writeAttribute('fo:font-size', $oTitle->getFont()->getSize() . 'pt');
        $this->xmlContent->writeAttribute('fo:font-style', $oTitle->getFont()->isItalic() ? 'italic' : 'normal');
        // > style:text-properties
        $this->xmlContent->endElement();
        // > style:style
        $this->xmlContent->endElement();
    }

    protected function writeWall(): void
    {
        // chart:wall
        $this->xmlContent->startElement('chart:wall');
        $this->xmlContent->writeAttribute('chart:style-name', 'styleWall');
        $this->xmlContent->endElement();
    }

    protected function writeWallStyle(Chart $chart): void
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
