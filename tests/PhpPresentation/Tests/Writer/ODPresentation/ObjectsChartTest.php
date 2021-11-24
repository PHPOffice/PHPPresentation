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
 * @copyright   2009-2015 PHPPresentation contributors
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
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\ODPresentation\ObjectsChart
 */
class ObjectsChartTest extends PhpPresentationTestCase
{
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
        $oShape->getPlotArea()->getAxisX()->getFont()->getColor()->setRGB('AABBCC');
        $oShape->getPlotArea()->getAxisX()->getFont()->setItalic(true);

        $oShape->getPlotArea()->getAxisY()->getFont()->getColor()->setRGB('00FF00');
        $oShape->getPlotArea()->getAxisY()->getFont()->setSize(16);
        $oShape->getPlotArea()->getAxisY()->getFont()->setName('Arial');

        $this->assertZipFileExists('Object 1/content.xml');

        $element = '/office:document-content/office:body/office:chart/chart:chart';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:class', 'chart:bar');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:text-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:color', '#AABBCC');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-style', 'italic');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-size', '10pt');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-family', 'Calibri');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:text-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:color', '#00FF00');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-style', 'normal');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-size', '16pt');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'fo:font-family', 'Arial');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testAxisTitleRotation(): void
    {
        $oSeries = new Series('Series', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipFileExists('Object 1/content.xml');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisXTitle\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'style:rotation-angle');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $value = rand(1, 360);
        $oShape->getPlotArea()->getAxisX()->setTitleRotation($value);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisXTitle\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'style:rotation-angle');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'style:rotation-angle', '-' . $value);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

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

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testChartDisplayBlankAs(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);
        $oShape->setDisplayBlankAs(Chart::BLANKAS_ZERO);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:treat-empty-cells');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:treat-empty-cells', 'use-zero');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $this->resetPresentationFile();
        $oShape->setDisplayBlankAs(Chart::BLANKAS_SPAN);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:treat-empty-cells');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:treat-empty-cells', 'ignore');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $this->resetPresentationFile();
        $oShape->setDisplayBlankAs(Chart::BLANKAS_GAP);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:treat-empty-cells');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:treat-empty-cells', 'leave-gap');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testLegend(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->setShowSeriesName(true);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oLine);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleLegend\']/style:chart-properties';
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
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_RIGHT);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'end');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_LEFT);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'start');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_BOTTOM);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'bottom');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_TOP);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'top');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oChart->getLegend()->setPosition(Legend::POSITION_TOPRIGHT);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:body/office:chart/chart:chart/chart:legend';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:legend-position', 'top-end');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

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
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

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
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testSeriesShowConfig(): void
    {
        $oSeries = new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oPie = new Pie();
        $oPie->addSeries($oSeries);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oPie);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\']/style:chart-properties';

        // $showCategoryName = false / $showPercentage = false / $showValue = true
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->setShowValue(false);
        $this->resetPresentationFile();

        // $showCategoryName = false / $showPercentage = false / $showValue = false
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        // $showCategoryName = false / $showPercentage = true / $showValue = true
        $oSeries->setShowValue(true);
        $oSeries->setShowPercentage(true);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value-and-percentage');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        // $showCategoryName = false / $showPercentage = true / $showValue = false
        $oSeries->setShowValue(false);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-number');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'percentage');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:data-label-text');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        // $showCategoryName = false / $showPercentage = true / $showValue = false
        $oSeries->setShowCategoryName(true);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:data-label-text');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-text', 'true');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTitleVisibility(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        $elementTitle = '/office:document-content/office:body/office:chart/chart:chart/chart:title';
        $elementStyle = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleTitle\']';

        $this->assertTrue($oShape->getTitle()->isVisible());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(true));
        $this->assertZipXmlElementExists('Object 1/content.xml', $elementTitle);
        $this->assertZipXmlElementExists('Object 1/content.xml', $elementStyle);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(false));
        $this->resetPresentationFile();
        $this->assertZipXmlElementNotExists('Object 1/content.xml', $elementTitle);
        $this->assertZipXmlElementNotExists('Object 1/content.xml', $elementStyle);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\']/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'draw:fill');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'draw:fill-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:fill-color', '#93A9CE');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeAxisBounds(): void
    {
        $value = mt_rand(0, 100);

        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:maximum');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oShape->getPlotArea()->getAxisX()->setMinBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:minimum', $value);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oShape->getPlotArea()->getAxisX()->setMinBounds(null);
        $oShape->getPlotArea()->getAxisX()->setMaxBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:maximum', $value);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oShape->getPlotArea()->getAxisX()->setMinBounds($value);
        $oShape->getPlotArea()->getAxisX()->setMaxBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:minimum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:minimum', $value);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:maximum');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:maximum', $value);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeAxisOutline(): void
    {
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:graphic-properties';

        $series = new Series('Series', $this->seriesData);
        $lineChart = new Line();
        $lineChart->addSeries($series);
        $shape = $this->oPresentation->getActiveSlide()->createChartShape();
        $shape->getPlotArea()->setType($lineChart);

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'draw:stroke');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'draw:stroke', 'none');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.035cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#000000');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

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

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeAxisTickLabelPosition(): void
    {
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';

        $oSeries = new Series('Series', $this->seriesData);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-label-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-label-position', 'near-axis');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:axis-position');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:tick-mark-position');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $this->resetPresentationFile();
        $oShape->getPlotArea()->getAxisX()->setTickLabelPosition(Axis::TICK_LABEL_POSITION_HIGH);

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-label-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-label-position', 'outside-end');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-position', '0');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:tick-mark-position');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $this->resetPresentationFile();
        $oShape->getPlotArea()->getAxisX()->setTickLabelPosition(Axis::TICK_LABEL_POSITION_LOW);

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-label-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-label-position', 'outside-start');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:axis-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:axis-position', '0');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:tick-mark-position');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-mark-position', 'at-axis');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeAxisUnit(): void
    {
        $value = mt_rand(0, 100);

        $series = new Series('Downloads', $this->seriesData);
        $line = new Line();
        $line->addSeries($series);
        $shape = $this->oPresentation->getActiveSlide()->createChartShape();
        $shape->getPlotArea()->setType($line);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interval-minor-divisor');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interval-major');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $shape->getPlotArea()->getAxisX()->setMinorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interval-major');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interval-minor-divisor');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interval-minor-divisor', $value);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $shape->getPlotArea()->getAxisX()->setMinorUnit(null);
        $shape->getPlotArea()->getAxisX()->setMajorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interval-minor-divisor');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interval-major');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interval-major', $value);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $shape->getPlotArea()->getAxisX()->setMinorUnit($value);
        $shape->getPlotArea()->getAxisX()->setMajorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interval-minor-divisor');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interval-minor-divisor', $value);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interval-major');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interval-major', $value);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'false');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '0');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:percentage');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeBarGroupingStacked(): void
    {
        $oBar = new Bar();
        $oBar->addSeries(new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']));
        $oBar->setBarGrouping(Bar::GROUPING_STACKED);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '100');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:percentage');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'value');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeBarGroupingPercentStacked(): void
    {
        $oBar = new Bar();
        $oBar->addSeries(new Series('Series', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']));
        $oBar->setBarGrouping(Bar::GROUPING_PERCENTSTACKED);
        $oChart = $this->oPresentation->getActiveSlide()->createChartShape();
        $oChart->getPlotArea()->setType($oBar);

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:stacked', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:overlap', '100');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:percentage', 'true');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:data-label-number', 'percentage');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'true');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:right-angled-axes');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'false');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:three-dimensional', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:right-angled-axes', 'true');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:vertical', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:three-dimensional');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:three-dimensional', 'true');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:right-angled-axes');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:right-angled-axes', 'true');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        // $oDoughnut->setHoleSize($randHoleSize);
        // $this->resetPresentationFile();

        $oSeries->setSeparator($randSeparator);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:automatic-styles/style:style/style:chart-properties/chart:label-separator/text:p';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlElementEquals('Object 1/content.xml', $element, $randSeparator);
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-inner', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-outer', 'false');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.035cm');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#000000');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-inner', 'false');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:tick-marks-major-outer', 'false');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:graphic-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-width', '0.035cm');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'svg:stroke-color', '#000000');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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
            $expectedElementStyle = '/office:document-content/office:automatic-styles/style:style[@style:name=\'' . $arrayTest['styleName'] . '\']/style:graphic-properties';

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
            $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElementGrid, 'chart:class');
            $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElementGrid, 'chart:class', $arrayTest['styleClass']);

            $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElementStyle);
            $this->assertZipXmlAttributeStartsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-width', $expectedSizeCm);
            $this->assertZipXmlAttributeEndsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-width', 'cm');
            $this->assertZipXmlAttributeStartsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-color', '#');
            $this->assertZipXmlAttributeEndsWith('Object 1/content.xml', $expectedElementStyle, 'svg:stroke-color', $expectedColor->getRGB());

            // chart:title : Element chart failed to validate attributes
            $this->assertIsSchemaOpenDocumentNotValid('1.2');
        }
    }

    public function testTypeLineMarker(): void
    {
        $expectedSymbol1 = Marker::SYMBOL_PLUS;
        $expectedSymbol2 = Marker::SYMBOL_DASH;
        $expectedSymbol3 = Marker::SYMBOL_DOT;
        $expectedSymbol4 = Marker::SYMBOL_TRIANGLE;
        $expectedSymbol5 = Marker::SYMBOL_NONE;

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:chart-properties';
        $expectedSize = mt_rand(1, 100);
        $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSize), 2, '.', '') . 'cm';

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getMarker()->setSymbol($expectedSymbol1)->setSize($expectedSize);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', $expectedSymbol1);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-width', $expectedSizeCm);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-height', $expectedSizeCm);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol2);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'horizontal-bar');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol3);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'circle');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol4);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'arrow-up');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol5);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-name');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-width');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-height');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeLineSeriesOutline(): void
    {
        $expectedWidth = mt_rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pointsToCentimeters($expectedWidth), 3, '.', '') . 'cm';

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:graphic-properties';

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

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', '0.079cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#4a7ebb');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->setOutline($oOutline);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', $expectedWidthCm);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#' . $oColor->getRGB());
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeLineSmooth(): void
    {
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';

        $oSeries = new Series('Downloads', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oLine->setIsSmooth(false);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interpolation');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $this->resetPresentationFile();
        $oLine->setIsSmooth(true);
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interpolation');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interpolation', 'cubic-spline');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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

        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:pie-offset', $value);
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisX\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:reverse-direction');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:reverse-direction', 'true');
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleAxisY\']/style:chart-properties';
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:reverse-direction');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeRadarSeriesOutline(): void
    {
        $expectedWidth = mt_rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pointsToCentimeters($expectedWidth), 3, '.', '') . 'cm';

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:graphic-properties';

        $color = new Color(Color::COLOR_YELLOW);

        $outline = new Outline();
        $outline->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($color);
        $outline->setWidth($expectedWidth); // (in points)

        $series = new Series('Downloads', $this->seriesData);

        $radarChart = new Radar();
        $radarChart->addSeries($series);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($radarChart);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', '0.079cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#4a7ebb');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $series->setOutline($outline);
        $radarChart->setSeries([$series]);
        $this->resetPresentationFile();

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', $expectedWidthCm);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#' . $color->getRGB());
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
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
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeScatterMarker(): void
    {
        $expectedSymbol1 = Marker::SYMBOL_PLUS;
        $expectedSymbol2 = Marker::SYMBOL_DASH;
        $expectedSymbol3 = Marker::SYMBOL_DOT;
        $expectedSymbol4 = Marker::SYMBOL_TRIANGLE;
        $expectedSymbol5 = Marker::SYMBOL_NONE;
        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:chart-properties';
        $expectedSize = mt_rand(1, 100);
        $expectedSizeCm = number_format(CommonDrawing::pointsToCentimeters($expectedSize), 2, '.', '') . 'cm';

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getMarker()->setSymbol($expectedSymbol1)->setSize($expectedSize);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', $expectedSymbol1);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-width', $expectedSizeCm);
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-height', $expectedSizeCm);

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol2);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'horizontal-bar');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol3);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'circle');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol4);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'chart:symbol-name', 'arrow-up');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->getMarker()->setSymbol($expectedSymbol5);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-name');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-width');
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $expectedElement, 'chart:symbol-height');

        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeScatterSeriesOutline(): void
    {
        $expectedWidth = mt_rand(1, 100);
        $expectedWidthCm = number_format(CommonDrawing::pointsToCentimeters($expectedWidth), 3, '.', '') . 'cm';

        $expectedElement = '/office:document-content/office:automatic-styles/style:style[@style:name=\'styleSeries0\'][@style:family=\'chart\']/style:graphic-properties';

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

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', '0.079cm');
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#4a7ebb');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oSeries->setOutline($oOutline);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $expectedElement);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-width');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-width', $expectedWidthCm);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $expectedElement, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $expectedElement, 'svg:stroke-color', '#' . $oColor->getRGB());
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTypeScatterSmooth(): void
    {
        $element = '/office:document-content/office:automatic-styles/style:style[@style:name=\'stylePlotArea\']/style:chart-properties';

        $oSeries = new Series('Downloads', $this->seriesData);

        $scatter = new Scatter();
        $scatter->addSeries($oSeries);
        $scatter->setIsSmooth(false);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($scatter);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeNotExists('Object 1/content.xml', $element, 'chart:interpolation');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $this->resetPresentationFile();
        $scatter->setIsSmooth(true);
        $oShape->getPlotArea()->setType($scatter);

        $this->assertZipFileExists('Object 1/content.xml');
        $this->assertZipXmlElementExists('Object 1/content.xml', $element);
        $this->assertZipXmlAttributeExists('Object 1/content.xml', $element, 'chart:interpolation');
        $this->assertZipXmlAttributeEquals('Object 1/content.xml', $element, 'chart:interpolation', 'cubic-spline');
        // chart:title : Element chart failed to validate attributes
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }
}
