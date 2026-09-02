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

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Chart\Legend;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Doughnut;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Radar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PHPUnit\Framework\Attributes\DataProvider;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\ODPresentation\ObjectsChart
 */
class ObjectsChartTest extends PhpPresentationTestCase
{
    private const CHART = '/office:document-content/office:body/office:chart/chart:chart';

    protected $writerName = 'ODPresentation';

    /**
     * @var array<string, string>
     */
    protected $seriesData = [
        'A' => '1',
        'B' => '2',
        'C' => '4',
        'D' => '3',
        'E' => '2',
    ];

    public function testAxisFont(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oSeries = new Series('Series', $this->seriesData);
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->getTickLabelFont()->getColor()->setRGB('AABBCC');
        $oShape->getPlotArea()->getAxisX()->getTickLabelFont()->setItalic(true);

        $oShape->getPlotArea()->getAxisY()->getTickLabelFont()->getColor()->setRGB('00FF00');
        $oShape->getPlotArea()->getAxisY()->getTickLabelFont()->setSize(16);
        $oShape->getPlotArea()->getAxisY()->getTickLabelFont()->setName('Arial');

        // the axis title is styled by the other font, and must not follow the labels
        $oShape->getPlotArea()->getAxisX()->getFont()->getColor()->setRGB('112233');
        $oShape->getPlotArea()->getAxisX()->getFont()->setSize(20);

        $this->assertZipFileExists('Object 1/content.xml');

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = $this->getAxisStyleXPath('x') . '/style:text-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:color', '#AABBCC');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-style', 'italic');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-size', '10pt');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-family', 'Calibri');

        $element = $this->getAxisStyleXPath('y') . '/style:text-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:color', '#00FF00');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-style', 'normal');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-size', '16pt');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-family', 'Arial');

        $element = $this->getAxisTitleStyleXPath('x') . '/style:text-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:color', '#112233');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-size', '20pt');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-style', 'normal');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testAxisTitleRotation(): void
    {
        $oSeries = new Series('Series', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipFileExists('Object 1/content.xml');

        $element = $this->getAxisTitleStyleXPath('x') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'style:rotation-angle');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $value = mt_rand(1, 360);
        $oShape->getPlotArea()->getAxisX()->setTitleRotation($value);
        $this->resetPresentationFile();

        $element = $this->getAxisTitleStyleXPath('x') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'style:rotation-angle');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:rotation-angle', '-' . $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testAxisVisibility(): void
    {
        $oSeries = new Series('Series', $this->seriesData);

        $oBar = new Bar();
        $oBar->addSeries($oSeries);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->setTitle('Axis X');
        $oShape->getPlotArea()->getAxisY()->setTitle('Axis Y');

        $oShape->getPlotArea()->getAxisX()->setIsVisible(false);
        $oShape->getPlotArea()->getAxisY()->setIsVisible(false);

        $this->assertZipFileExists('Object 1/content.xml');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\'x\']/chart:title';

        $this->assertZipXmlElementNotExists('Object 1/content.xml', $element);

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\'y\']/chart:title';

        $this->assertZipXmlElementNotExists('Object 1/content.xml', $element);

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oShape->getPlotArea()->getAxisX()->setIsVisible(true);
        $oShape->getPlotArea()->getAxisY()->setIsVisible(true);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\'x\']/chart:title';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:style-name', 'styleAxisXTitle');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\'x\']/chart:title/text:p';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlElementEquals('Object 1/content.xml', $element, 'Axis X');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\'y\']/chart:title';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:style-name', 'styleAxisYTitle');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\'y\']/chart:title/text:p';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlElementEquals('Object 1/content.xml', $element, 'Axis Y');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testChartDisplayBlankAs(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);
        $oShape->setDisplayBlankAs(Chart::BLANKAS_ZERO);

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';
        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:treat-empty-cells');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:treat-empty-cells', 'use-zero');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $oShape->setDisplayBlankAs(Chart::BLANKAS_SPAN);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:treat-empty-cells');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:treat-empty-cells', 'ignore');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $oShape->setDisplayBlankAs(Chart::BLANKAS_GAP);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:treat-empty-cells');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:treat-empty-cells', 'leave-gap');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testLegend(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);

        $element = $this->getLegendStyleXPath() . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:auto-position', 'true');
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'end');
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row/table:table-cell';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-header-rows/table:table-row/table:table-cell[@office:value-type=\'string\']';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_RIGHT);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'end');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_LEFT);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'start');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_BOTTOM);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'bottom');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_TOP);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'top');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_TOPRIGHT);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'top-end');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testLegendVisibility(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oChart->getLegend()->setVisible(false);
        $this->resetPresentationFile();

        $this->assertZipXmlElementNotExists('Object 1/content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testChartFill(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);

        $element = $this->getChartStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill', 'none');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oChart->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill-color', '#FFFFFF');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testChartFillSetBackToNull(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);
        $oChart->setFill(null);

        // A chart is born at `FILL_NONE`, so a chart that named no fill is drawn like one that
        // refused it. Reading the fill used to be enough to kill the writer outright.
        $element = $this->getChartStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill', 'none');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:stroke', 'none');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testSeriesFill(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = $this->getSeriesStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'draw:fill');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'draw:fill-color');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4472C4'));
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill-color', '#4472C4');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testPlotAreaDataLabels(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // A serie showing nothing must not leave a default label on the plot area
        $oSeries->setShowValue(false);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->setShowPercentage(true);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'percentage');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testSeriesValues(): void
    {
        $series = new Series('Series', ['Jan' => null]);

        $pie = new Pie();
        $pie->addSeries($series);

        $chart = $this->oPresentation->getActiveSlide()->createChartShape();
        $chart->getPlotArea()->setType($pie);

        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-rows/table:table-row/table:table-cell[2]';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'office:value-type');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'office:value-type', 'float');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'office:value');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'office:value', 'NaN');

        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-rows/table:table-row/table:table-cell[2]/text:p';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlElementEquals('Object 1/content.xml', $element, 'NaN');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();

        $series = new Series('Series', ['Jan' => '12.3']);
        $chart->getPlotArea()->getType()->setSeries([$series]);

        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-rows/table:table-row/table:table-cell[2]';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'office:value-type');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'office:value-type', 'float');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'office:value');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'office:value', '12.3');

        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-rows/table:table-row/table:table-cell[2]/text:p';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlElementEquals('Object 1/content.xml', $element, '12.3');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();

        $series = new Series('Series', ['Jan' => 'data']);
        $chart->getPlotArea()->getType()->setSeries([$series]);

        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-rows/table:table-row/table:table-cell[2]';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'office:value-type');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'office:value-type', 'string');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'office:value');

        $element = '/office:document-content/office:body/office:chart/chart:chart/table:table/table:table-rows/table:table-row/table:table-cell[2]/text:p';

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlElementEquals('Object 1/content.xml', $element, 'data');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testSeriesShowConfig(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oPie = new Pie();
        $oPie->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie);

        $element = $this->getSeriesStyleXPath() . '/style:chart-properties';

        // $showCategoryName = false / $showPercentage = false / $showValue = true
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->setShowValue(false);
        $this->resetPresentationFile();

        // $showCategoryName = false / $showPercentage = false / $showValue = false
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // $showCategoryName = false / $showPercentage = true / $showValue = true
        $oSeries->setShowValue(true);
        $oSeries->setShowPercentage(true);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value-and-percentage');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // $showCategoryName = false / $showPercentage = true / $showValue = false
        $oSeries->setShowValue(false);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'percentage');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // $showCategoryName = false / $showPercentage = true / $showValue = false
        $oSeries->setShowCategoryName(true);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-text');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-text', 'true');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTitleVisibility(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        $elementTitle = '/office:document-content/office:body/office:chart/chart:chart/chart:title';
        $elementStyle = $this->getTitleStyleXPath();

        self::assertTrue($oShape->getTitle()->isVisible());
        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(true));
        $this->assertZipXmlElementExists('Object 1/content.xml', $elementTitle);
        $this->assertZipXmlElementExists('Object 1/content.xml', $elementStyle);

        $this->assertIsSchemaOpenDocumentValid('1.2');

        self::assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(false));
        $this->resetPresentationFile();
        $this->assertZipXmlElementNotExists('Object 1/content.xml', $elementTitle);
        $this->assertZipXmlElementNotExists('Object 1/content.xml', $elementStyle);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeArea(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->getFill()->setStartColor(new Color('FF93A9CE'));
        $oArea = new Area();
        $oArea->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oArea);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:area');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);

        $element = $this->getSeriesStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'draw:fill');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'draw:fill-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill-color', '#93A9CE');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeAxisBounds(): void
    {
        $value = mt_rand(0, 100);

        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $element = $this->getAxisStyleXPath('x') . '/style:chart-properties';

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:maximum');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oShape->getPlotArea()->getAxisX()->setMinBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:minimum', $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oShape->getPlotArea()->getAxisX()->setMinBounds(null);
        $oShape->getPlotArea()->getAxisX()->setMaxBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:maximum', $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oShape->getPlotArea()->getAxisX()->setMinBounds($value);
        $oShape->getPlotArea()->getAxisX()->setMaxBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:minimum', $value);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:maximum', $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeAxisOutline(): void
    {
        $series = new Series('Series', $this->seriesData);
        $lineChart = new Line();
        $lineChart->addSeries($series);
        $shape = $this->oPresentation->getActiveSlide()->createChartShape();
        $shape->getPlotArea()->setType($lineChart);

        $element = $this->getAxisStyleXPath('x') . '/style:graphic-properties';

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'draw:stroke');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:stroke', 'none');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.035cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#000000');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $shape->getPlotArea()->getAxisX()->getOutline()->setWidth(10);
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID);
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->getStartColor()->setRGB('ABCDEF');

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'draw:stroke');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:stroke', 'solid');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.353cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#ABCDEF');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeAxisTickLabelPosition(): void
    {
        $oSeries = new Series('Series', $this->seriesData);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $element = $this->getAxisStyleXPath('x') . '/style:chart-properties';

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-label-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-label-position', 'near-axis');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:axis-position');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:tick-mark-position');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $oShape->getPlotArea()->getAxisX()->setTickLabelPosition(Axis::TICK_LABEL_POSITION_HIGH);

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-label-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-label-position', 'outside-end');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-position', '0');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:tick-mark-position');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $oShape->getPlotArea()->getAxisX()->setTickLabelPosition(Axis::TICK_LABEL_POSITION_LOW);

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-label-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-label-position', 'outside-start');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-position', '0');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:tick-mark-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-mark-position', 'at-axis');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeAxisUnit(): void
    {
        $value = max(1, mt_rand(0, 100));

        $series = new Series('Downloads', $this->seriesData);
        $line = new Line();
        $line->addSeries($series);
        $shape = $this->oPresentation->getActiveSlide()->createChartShape();
        $shape->getPlotArea()->setType($line);

        $element = $this->getAxisStyleXPath('x') . '/style:chart-properties';

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interval-minor-divisor');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interval-major');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $shape->getPlotArea()->getAxisX()->setMinorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interval-major');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interval-minor-divisor');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interval-minor-divisor', $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $shape->getPlotArea()->getAxisX()->setMinorUnit(null);
        $shape->getPlotArea()->getAxisX()->setMajorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interval-minor-divisor');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interval-major');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interval-major', $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $shape->getPlotArea()->getAxisX()->setMinorUnit($value);
        $shape->getPlotArea()->getAxisX()->setMajorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interval-minor-divisor');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interval-minor-divisor', $value);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interval-major');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interval-major', $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeBar(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'false');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '0');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:percentage');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeBarGroupingStacked(): void
    {
        $oBar = new Bar();
        $oBar->addSeries(new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']));
        $oBar->setBarGrouping(Bar::GROUPING_STACKED);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '100');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:percentage');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeBarGroupingPercentStacked(): void
    {
        $oBar = new Bar();
        $oBar->addSeries(new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']));
        $oBar->setBarGrouping(Bar::GROUPING_PERCENTSTACKED);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '100');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:percentage', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'percentage');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeBarDataPointsRepeated(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2', 'Apr' => '4']);
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        // The data points describe the four values of the serie, no more and no less
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point[1]';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:repeated', 1);
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point[2]';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:style-name', 'styleSeries0_1');
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point[3]';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:repeated', 2);
        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point[4]';
        $this->assertZipXmlElementNotExists('Object 1/content.xml', $element);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testDataPointOutlines(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        // A slice with a white border of its own
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointOutline(0)->setWidth(2)->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));
        // A slice made invisible : no fill and no border
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_NONE);
        $oSeries->getDataPointOutline(1)->getFill()->setFillType(Fill::FILL_NONE);
        // A slice carrying an outline only
        $oSeries->getDataPointOutline(2)->setWidth(1)->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oDoughnut = new Doughnut();
        $oDoughnut->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oDoughnut);

        $element = $this->getDataPointStyleXPath(1) . '/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:stroke', 'solid');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#FFFFFF');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'svg:stroke-width');

        $element = $this->getDataPointStyleXPath(2) . '/style:graphic-properties';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:stroke', 'none');

        $element = $this->getDataPointStyleXPath(3) . '/style:graphic-properties';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:stroke', 'solid');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#AB4744');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'draw:fill');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeBarHorizontal(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
        $oBar = new Bar();
        $oBar->setBarDirection(Bar::DIRECTION_HORIZONTAL);
        $oBar->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'true');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:right-angled-axes');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeBar3D(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
        $oBar3D = new Bar3D();
        $oBar3D->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar3D);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'false');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:three-dimensional', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:right-angled-axes', 'true');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeBar3DHorizontal(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
        $oBar3D = new Bar3D();
        $oBar3D->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);
        $oBar3D->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar3D);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:three-dimensional', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:right-angled-axes', 'true');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeDoughnut(): void
    {
        // $randHoleSize = mt_rand(10, 90);
        $randSeparator = chr(mt_rand(ord('A'), ord('Z')));

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oDoughnut = new Doughnut();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oDoughnut->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oDoughnut);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:ring');
        $element = '/office:document-content/office:automatic-styles/style:style/style:chart-properties/chart:label-separator/text:p';
        $this->assertZipXmlElementNotExists('Object 1/content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // $oDoughnut->setHoleSize($randHoleSize);
        // $this->resetPresentationFile();

        $oSeries->setSeparator($randSeparator);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:automatic-styles/style:style/style:chart-properties/chart:label-separator/text:p';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlElementEquals('Object 1/content.xml', $element, $randSeparator);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeLine(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:line');

        $element = $this->getAxisStyleXPath('x') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-inner', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-outer', 'false');

        $element = $this->getAxisStyleXPath('x') . '/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.035cm');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#000000');

        $element = $this->getAxisStyleXPath('y') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-inner', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-outer', 'false');

        $element = $this->getAxisStyleXPath('y') . '/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.035cm');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#000000');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeLineGridlines(): void
    {
        $arrayTests = [
            [
                'dimension' => 'x',
                'styleName' => 'styleAxisXGridlinesMajor',
                'styleClass' => 'major',
                'methodAxis' => 'getAxisX',
                'methodGrid' => 'setMajorGridlines',
            ],
            [
                'dimension' => 'x',
                'styleName' => 'styleAxisXGridlinesMinor',
                'styleClass' => 'minor',
                'methodAxis' => 'getAxisX',
                'methodGrid' => 'setMinorGridlines',
            ],
            [
                'dimension' => 'y',
                'styleName' => 'styleAxisYGridlinesMajor',
                'styleClass' => 'major',
                'methodAxis' => 'getAxisY',
                'methodGrid' => 'setMajorGridlines',
            ],
            [
                'dimension' => 'y',
                'styleName' => 'styleAxisYGridlinesMinor',
                'styleClass' => 'minor',
                'methodAxis' => 'getAxisY',
                'methodGrid' => 'setMinorGridlines',
            ],
        ];
        $expectedColor = new Color(Color::COLOR_BLUE);

        foreach ($arrayTests as $arrayTest) {
            $this->resetPresentationFile();
            $this->oPresentation->removeSlideByIndex(0)->createSlide();

            $expectedSizePts = mt_rand(1, 100);
            $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSizePts), 2, '.', '') . 'cm';
            $expectedElementGrid = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:axis[@chart:dimension=\'' . $arrayTest['dimension'] . '\']/chart:grid';

            $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
            $oLine = new Line();
            $oLine->addSeries(new Series('Downloads', $this->seriesData));
            $oShape->getPlotArea()->setType($oLine);
            $oGridlines = new Gridlines();
            $oGridlines->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($expectedColor);
            $oGridlines->getOutline()->setWidth($expectedSizePts);
            $oShape->getPlotArea()->{$arrayTest['methodAxis']}()->{$arrayTest['methodGrid']}($oGridlines);

            $this->assertZipFileExists('Object 1/content.xml');
            $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElementGrid);
            $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElementGrid, 'chart:style-name');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElementGrid, 'chart:style-name', $arrayTest['styleName']);
            $expectedElementStyle = $this->getGridlinesStyleXPath($arrayTest['dimension'], $arrayTest['styleClass']) . '/style:graphic-properties';
            $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElementGrid, 'chart:class');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElementGrid, 'chart:class', $arrayTest['styleClass']);

            $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElementStyle);
            $this->assertZipXmlAttributeStartsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-width', $expectedSizeCm);
            $this->assertZipXmlAttributeEndsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-width', 'cm');
            $this->assertZipXmlAttributeStartsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-color', '#');
            $this->assertZipXmlAttributeEndsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-color', $expectedColor->getRGB());

            $this->assertIsSchemaOpenDocumentValid('1.2');
        }
    }

    public function testTypeLineMarker(): void
    {
        $expectedSymbol1 = Marker::SYMBOL_PLUS;
        $expectedSymbol2 = Marker::SYMBOL_DASH;
        $expectedSymbol3 = Marker::SYMBOL_DOT;
        $expectedSymbol4 = Marker::SYMBOL_TRIANGLE;
        $expectedSymbol5 = Marker::SYMBOL_NONE;

        $expectedSize = mt_rand(1, 100);
        $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSize), 2, '.', '') . 'cm';

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getMarker()->setSymbol($expectedSymbol1)->setSize($expectedSize);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $expectedElement = $this->getSeriesStyleXPath() . '[@style:family=\'chart\']/style:chart-properties';

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', $expectedSymbol1);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-width', $expectedSizeCm);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-height', $expectedSizeCm);

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol2);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'horizontal-bar');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol3);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'circle');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol4);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'arrow-up');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol5);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-name');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-width');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-height');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeLineSeriesOutline(): void
    {
        $expectedWidth = mt_rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pixelsToCentimeters($expectedWidth), 3, '.', '') . 'cm';

        $oColor = new Color(Color::COLOR_YELLOW);
        $oOutline = new Outline();
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor($oColor);
        $oOutline->setWidth($expectedWidth); // (in points)

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $expectedElement = $this->getSeriesStyleXPath() . '[@style:family=\'chart\']/style:graphic-properties';

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', '0.079cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#4a7ebb');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->setOutline($oOutline);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', $expectedWidthCm);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#' . $oColor->getRGB());
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeLineSmooth(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oLine->setIsSmooth(false);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interpolation');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $oLine->setIsSmooth(true);
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interpolation');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interpolation', 'cubic-spline');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypePie(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
        $oPie = new Pie();
        $oPie->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:circle');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);

        $element = $this->getAxisStyleXPath('x') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        $element = $this->getAxisStyleXPath('y') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypePie3D(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
        $oPie3D = new Pie3D();
        $oPie3D->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie3D);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:circle');

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:plot-area/chart:series/chart:data-point';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);

        $element = $this->getAxisStyleXPath('x') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        $element = $this->getAxisStyleXPath('y') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypePie3DExplosion(): void
    {
        $value = mt_rand(0, 100);

        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oPie3D = new Pie3D();
        $oPie3D->setExplosion($value);
        $oPie3D->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie3D);

        $element = $this->getSeriesStyleXPath() . '[@style:family=\'chart\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:pie-offset', $value);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeRadar(): void
    {
        $series = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $series->setShowSeriesName(true);

        $radarChart = new Radar();
        $radarChart->addSeries($series);

        $chart = $this->oPresentation->getActiveSlide()->createChartShape();
        $chart->getPlotArea()->setType($radarChart);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:class');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:radar');
        $element = $this->getAxisStyleXPath('x') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:reverse-direction');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');
        $element = $this->getAxisStyleXPath('y') . '/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:reverse-direction');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeRadarSeriesOutline(): void
    {
        $expectedWidth = mt_rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pixelsToCentimeters($expectedWidth), 3, '.', '') . 'cm';

        $color = new Color(Color::COLOR_YELLOW);

        $outline = new Outline();
        $outline->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($color);
        $outline->setWidth($expectedWidth); // (in points)

        $series = new Series('Downloads', $this->seriesData);

        $radarChart = new Radar();
        $radarChart->addSeries($series);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($radarChart);

        $expectedElement = $this->getSeriesStyleXPath() . '[@style:family=\'chart\']/style:graphic-properties';

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', '0.079cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#4a7ebb');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $series->setOutline($outline);
        $radarChart->setSeries([$series]);
        $this->resetPresentationFile();

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', $expectedWidthCm);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#' . $color->getRGB());
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeScatter(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oScatter = new Scatter();
        $oScatter->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oScatter);

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:scatter');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeScatterMarker(): void
    {
        $expectedSymbol1 = Marker::SYMBOL_PLUS;
        $expectedSymbol2 = Marker::SYMBOL_DASH;
        $expectedSymbol3 = Marker::SYMBOL_DOT;
        $expectedSymbol4 = Marker::SYMBOL_TRIANGLE;
        $expectedSymbol5 = Marker::SYMBOL_NONE;
        $expectedSize = mt_rand(1, 100);
        $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSize), 2, '.', '') . 'cm';

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getMarker()->setSymbol($expectedSymbol1)->setSize($expectedSize);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $expectedElement = $this->getSeriesStyleXPath() . '[@style:family=\'chart\']/style:chart-properties';

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', $expectedSymbol1);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-width', $expectedSizeCm);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-height', $expectedSizeCm);

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol2);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'horizontal-bar');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol3);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'circle');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol4);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'arrow-up');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol5);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-name');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-width');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-height');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeScatterSeriesOutline(): void
    {
        $expectedWidth = mt_rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pixelsToCentimeters($expectedWidth), 3, '.', '') . 'cm';

        $oColor = new Color(Color::COLOR_YELLOW);
        $oOutline = new Outline();
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($oColor);
        $oOutline->setWidth($expectedWidth); // (in points)

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $expectedElement = $this->getSeriesStyleXPath() . '[@style:family=\'chart\']/style:graphic-properties';

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', '0.079cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#4a7ebb');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSeries->setOutline($oOutline);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', $expectedWidthCm);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#' . $oColor->getRGB());
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTypeScatterSmooth(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);

        $scatter = new Scatter();
        $scatter->addSeries($oSeries);
        $scatter->setIsSmooth(false);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($scatter);

        $element = $this->getPlotAreaStyleXPath() . '/style:chart-properties';

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interpolation');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $scatter->setIsSmooth(true);
        $oShape->getPlotArea()->setType($scatter);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interpolation');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interpolation', 'cubic-spline');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @return array<array{0: string, 1: string, 2: string, 3: null|string}>
     */
    public static function dataProviderUnderlines(): array
    {
        return [
            [Font::UNDERLINE_DASH, 'dash', 'single', null],
            [Font::UNDERLINE_DASHHEAVY, 'dash', 'single', 'bold'],
            [Font::UNDERLINE_DASHLONG, 'long-dash', 'single', null],
            [Font::UNDERLINE_DASHLONGHEAVY, 'long-dash', 'single', 'bold'],
            [Font::UNDERLINE_DOTHASH, 'dot-dash', 'single', null],
            [Font::UNDERLINE_DOTHASHHEAVY, 'dot-dash', 'single', 'bold'],
            [Font::UNDERLINE_DOTDOTDASH, 'dot-dot-dash', 'single', null],
            [Font::UNDERLINE_DOTDOTDASHHEAVY, 'dot-dot-dash', 'single', 'bold'],
            [Font::UNDERLINE_DOTTED, 'dotted', 'single', null],
            [Font::UNDERLINE_DOTTEDHEAVY, 'dotted', 'single', 'bold'],
            [Font::UNDERLINE_DOUBLE, 'solid', 'double', null],
            [Font::UNDERLINE_HEAVY, 'solid', 'single', 'bold'],
            [Font::UNDERLINE_SINGLE, 'solid', 'single', null],
            [Font::UNDERLINE_WAVY, 'wave', 'single', null],
            [Font::UNDERLINE_WAVYDOUBLE, 'wave', 'double', null],
            [Font::UNDERLINE_WAVYHEAVY, 'wave', 'single', 'bold'],
            [Font::UNDERLINE_WORDS, 'solid', 'single', null],
        ];
    }

    /**
     * @dataProvider dataProviderUnderlines
     */
    #[DataProvider('dataProviderUnderlines')]
    public function testAxisFontUnderline(string $underline, string $style, string $type, ?string $width): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oSeries = new Series('Series', $this->seriesData);
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->getTickLabelFont()->setUnderline($underline);

        $element = $this->getAxisStyleXPath('x') . '/style:text-properties';
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:text-underline-style', $style);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:text-underline-type', $type);
        if (null === $width) {
            $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'style:text-underline-width');
        } else {
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:text-underline-width', $width);
        }
        // `words` is the one underline ODF spells with a mode rather than a style of its own
        if (Font::UNDERLINE_WORDS === $underline) {
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:text-underline-mode', 'skip-white-space');
        } else {
            $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'style:text-underline-mode');
        }

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testFontStateOfEveryChartStyle(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oSeries = new Series('Series', $this->seriesData);
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->setTitle('Axis');
        $oShape->getTitle()->setVisible(true);
        $oShape->getLegend()->setVisible(true);

        $fonts = [
            $oShape->getPlotArea()->getAxisX()->getFont(),
            $oShape->getPlotArea()->getAxisX()->getTickLabelFont(),
            $oShape->getLegend()->getFont(),
            $oSeries->getFont(),
            $oShape->getTitle()->getFont(),
        ];
        foreach ($fonts as $oFont) {
            $oFont->setBold(true);
            $oFont->setItalic(true);
            $oFont->setUnderline(Font::UNDERLINE_DOUBLE);
            $oFont->setStrikethrough(Font::STRIKE_DOUBLE);
        }

        // the axis has a font per half: the tick labels and the title it carries
        foreach ([
            $this->getAxisStyleXPath('x'),
            $this->getAxisTitleStyleXPath('x'),
            $this->getLegendStyleXPath(),
            $this->getSeriesStyleXPath(),
            $this->getTitleStyleXPath(),
        ] as $styleXPath) {
            $element = $styleXPath . '/style:text-properties';
            $this->assertZipXmlElementExists('Object 1/content.xml', $element);
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-weight', 'bold');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-style', 'italic');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:text-underline-style', 'solid');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:text-underline-type', 'double');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:text-line-through-style', 'solid');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:text-line-through-type', 'double');
        }

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testFontStateLeftAloneIsNotWritten(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oSeries = new Series('Series', $this->seriesData);
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);

        $element = $this->getAxisStyleXPath('x') . '/style:text-properties';
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'fo:font-weight');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'style:text-underline-style');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'style:text-line-through-style');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testFontSetBackToNull(): void
    {
        $oLine = new Line();
        $oLine->addSeries(new Series('Downloads', $this->seriesData));
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        // every one of these accepts null, and used to store it -- the writer then read a font
        // off it and died. Passing no argument at all is the same call
        $oShape->getPlotArea()->getAxisY()->setFont();
        $oShape->getPlotArea()->getAxisY()->setTickLabelFont();
        $oShape->getPlotArea()->getAxisX()->setFont(null);
        $oShape->getTitle()->setFont(null);
        $oShape->getLegend()->setFont(null);
        $oLine->getSeries()[0]->setFont(null);

        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('Alpha')->setFont(null);
        $oRichText->getActiveParagraph()->setFont(null);

        $element = $this->getAxisStyleXPath('y') . '/style:text-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-family', 'Calibri');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * The definition of the automatic style that $referenceXPath names.
     *
     * A chart style is addressed by the name the document carries, the way a consumer reads it,
     * rather than by one the test spells out: a definition nothing points at fails here.
     */
    private function getChartAutomaticStyleXPath(string $referenceXPath): string
    {
        $styleName = $this->getZipXmlAttributeValue('Object 1/content.xml', $referenceXPath, 'chart:style-name');

        return '/office:document-content/office:automatic-styles/style:style[@style:name=\'' . $styleName . '\']';
    }

    private function getChartStyleXPath(): string
    {
        return $this->getChartAutomaticStyleXPath(self::CHART);
    }

    private function getTitleStyleXPath(): string
    {
        return $this->getChartAutomaticStyleXPath(self::CHART . '/chart:title');
    }

    private function getLegendStyleXPath(): string
    {
        return $this->getChartAutomaticStyleXPath(self::CHART . '/chart:legend');
    }

    private function getPlotAreaStyleXPath(): string
    {
        return $this->getChartAutomaticStyleXPath(self::CHART . '/chart:plot-area');
    }

    private function getAxisStyleXPath(string $dimension): string
    {
        return $this->getChartAutomaticStyleXPath($this->getAxisXPath($dimension));
    }

    private function getAxisTitleStyleXPath(string $dimension): string
    {
        return $this->getChartAutomaticStyleXPath($this->getAxisXPath($dimension) . '/chart:title');
    }

    private function getSeriesStyleXPath(int $series = 1): string
    {
        return $this->getChartAutomaticStyleXPath($this->getSeriesXPath($series));
    }

    private function getDataPointStyleXPath(int $dataPoint, int $series = 1): string
    {
        return $this->getChartAutomaticStyleXPath(sprintf('%s/chart:data-point[%d]', $this->getSeriesXPath($series), $dataPoint));
    }

    private function getGridlinesStyleXPath(string $dimension, string $class): string
    {
        return $this->getChartAutomaticStyleXPath(sprintf(
            '%s/chart:grid[@chart:class=\'%s\']',
            $this->getAxisXPath($dimension),
            $class
        ));
    }

    private function getAxisXPath(string $dimension): string
    {
        return sprintf('%s/chart:plot-area/chart:axis[@chart:dimension=\'%s\']', self::CHART, $dimension);
    }

    private function getSeriesXPath(int $series): string
    {
        return sprintf('%s/chart:plot-area/chart:series[%d]', self::CHART, $series);
    }
}
