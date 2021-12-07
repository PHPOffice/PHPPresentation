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

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\Common\Drawing;
use PhpOffice\PhpPresentation\Exception\UndefinedChartTypeException;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType;
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
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptChartsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

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

    public function testChartDisplayBlankAs(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);
        $oShape->setDisplayBlankAs(Chart::BLANKAS_ZERO);

        $element = '/c:chartSpace/c:chart/c:dispBlanksAs';

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', Chart::BLANKAS_ZERO);

        $this->assertIsSchemaECMA376Valid();

        $this->resetPresentationFile();
        $oShape->setDisplayBlankAs(Chart::BLANKAS_SPAN);

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', Chart::BLANKAS_SPAN);

        $this->assertIsSchemaECMA376Valid();

        $this->resetPresentationFile();
        $oShape->setDisplayBlankAs(Chart::BLANKAS_GAP);

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', Chart::BLANKAS_GAP);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testChartIncludeSpreadsheet(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setIncludeSpreadsheet(true);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $this->assertTrue($oShape->hasIncludedSpreadsheet());

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipFileExists('ppt/embeddings/' . $oShape->getIndexedFilename() . '.xlsx');

        $this->assertZipFileExists('ppt/charts/_rels/' . $oShape->getIndexedFilename() . '.rels');
        $element = '/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"]';
        $this->assertZipXmlElementExists('ppt/charts/_rels/' . $oShape->getIndexedFilename() . '.rels', $element);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:strRef';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:strRef/c:f';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'Sheet1!$A$1');
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:strRef/c:strCache';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:strRef/c:strCache/c:pt';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:strRef/c:strCache/c:pt/c:v';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'Downloads');
        $element = '/c:chartSpace/c:externalData';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'r:id');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'r:id', 'rId1');
        $element = '/c:chartSpace/c:externalData/c:autoUpdate';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '0');

        $this->assertIsSchemaECMA376Valid();

        $this->resetPresentationFile();
        $oShape->setIncludeSpreadsheet(false);

        $this->assertFalse($oShape->hasIncludedSpreadsheet());

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipFileNotExists('ppt/charts/_rels/' . $oShape->getIndexedFilename() . '.rels');
        $this->assertZipFileNotExists('ppt/embeddings/' . $oShape->getIndexedFilename() . '.xlsx');

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'Downloads');
        $element = '/c:chartSpace/c:externalData';
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testPlotAreaBadType(): void
    {
        $this->expectException(UndefinedChartTypeException::class);
        $this->expectExceptionMessage('The chart type has not been defined');

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        /** @var AbstractType $stub */
        $stub = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType');
        $oShape->getPlotArea()->setType($stub);

        $this->writePresentationFile($this->oPresentation, 'PowerPoint2007');
    }

    public function testTitleVisibilityTrue(): void
    {
        $element = '/c:chartSpace/c:chart/c:autoTitleDeleted';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        // Default
        $this->assertTrue($oShape->getTitle()->isVisible());

        // Set Visible : TRUE
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(true));
        $this->assertTrue($oShape->getTitle()->isVisible());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '0');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testTitleVisibilityFalse(): void
    {
        $element = '/c:chartSpace/c:chart/c:autoTitleDeleted';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        // Set Visible : FALSE
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Title', $oShape->getTitle()->setVisible(false));
        $this->assertFalse($oShape->getTitle()->isVisible());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '1');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisBounds(): void
    {
        $value = mt_rand(0, 100);

        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFill()->setStartColor(new Color('FFAABBCC'));
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $elementMax = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:scaling/c:max';
        $elementMin = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:scaling/c:min';

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisX()->setMinBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin, 'val', $value);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisX()->setMinBounds(null);
        $oShape->getPlotArea()->getAxisX()->setMaxBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', $value);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisX()->setMinBounds($value);
        $oShape->getPlotArea()->getAxisX()->setMaxBounds($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin, 'val', $value);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', $value);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisCrosses(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFill()->setStartColor(new Color('FFAABBCC'));
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $elementCrosses = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:crosses';
        $elementCrossesAt = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:crossesAt';
        $elementAxPos = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:axPos';

        // Default autoZero
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrosses);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrosses, 'val', 'autoZero');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementAxPos, 'val', 'b');
        $this->assertIsSchemaECMA376Valid();

        // Crosses max
        $oShape->getPlotArea()->getAxisX()->setCrossesAt(Axis::CROSSES_MAX);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrosses);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrosses, 'val', 'max');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementAxPos, 'val', 't');
        $this->assertIsSchemaECMA376Valid();

        // Crosses min
        $oShape->getPlotArea()->getAxisX()->setCrossesAt(Axis::CROSSES_MIN);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrosses);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrosses, 'val', 'min');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementAxPos, 'val', 'b');
        $this->assertIsSchemaECMA376Valid();

        // Crosses custom value
        $oShape->getPlotArea()->getAxisX()->setCrossesAt('10');
        $this->resetPresentationFile();

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrosses);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrossesAt);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementCrossesAt, 'val', '10');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementAxPos, 'val', 'b');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testIsReversedOrder(): void
    {
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:scaling/c:orientation';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        // default
        $this->assertFalse($oShape->getPlotArea()->getAxisX()->isReversedOrder());
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 'minMax');
        $this->assertIsSchemaECMA376Valid();

        // reversed order
        $this->assertInstanceOf(Axis::class, $oShape->getPlotArea()->getAxisX()->setIsReversedOrder(true));
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 'maxMin');
        $this->assertIsSchemaECMA376Valid();

        // reset reversed order
        $this->assertInstanceOf(Axis::class, $oShape->getPlotArea()->getAxisX()->setIsReversedOrder(false));
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 'minMax');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisFont(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oBar = new Bar();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->getFont()->getColor()->setRGB('AABBCC');
        $oShape->getPlotArea()->getAxisX()->getFont()->setItalic(true);
        $oShape->getPlotArea()->getAxisX()->getFont()->setStrikethrough(true);
        $oShape->getPlotArea()->getAxisY()->getFont()->getColor()->setRGB('00FF00');
        $oShape->getPlotArea()->getAxisY()->getFont()->setSize(16);
        $oShape->getPlotArea()->getAxisY()->getFont()->setUnderline(Font::UNDERLINE_DASH);

        $pathShape = 'ppt/charts/' . $oShape->getIndexedFilename();
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:solidFill/a:srgbClr';
        $this->assertZipXmlElementExists($pathShape, $element);
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'val', 'AABBCC');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists($pathShape, $element);
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'i', 'true');
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'sz', 1000);
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'strike', 'sngStrike');
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'u', Font::UNDERLINE_NONE);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr/a:solidFill/a:srgbClr';
        $this->assertZipXmlElementExists($pathShape, $element);
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'val', '00FF00');

        $element = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:title/c:tx/c:rich/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists($pathShape, $element);
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'i', 'false');
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'sz', 1600);
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'strike', 'noStrike');
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'u', Font::UNDERLINE_DASH);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisOutline(): void
    {
        $expectedWidthX = 2;
        $expectedColorX = 'ABCDEF';
        $expectedWidthY = 4;
        $expectedColorY = '012345';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oBar = new Bar();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);
        $oShape->getPlotArea()->getAxisX()->getOutline()->setWidth($expectedWidthX);
        $oShape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID);
        $oShape->getPlotArea()->getAxisX()->getOutline()->getFill()->getStartColor()->setRGB($expectedColorX);
        $oShape->getPlotArea()->getAxisY()->getOutline()->setWidth($expectedWidthY);
        $oShape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID);
        $oShape->getPlotArea()->getAxisY()->getOutline()->getFill()->getStartColor()->setRGB($expectedColorY);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:spPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:spPr/a:ln';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'w', Drawing::pointsToEmu($expectedWidthX));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:spPr/a:ln/a:solidFill/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', $expectedColorX);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:spPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:spPr/a:ln';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'w', Drawing::pointsToEmu($expectedWidthY));
        $element = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:spPr/a:ln/a:solidFill/a:srgbClr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', $expectedColorY);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisTickLabelPosition(): void
    {
        $element = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:tickLblPos';

        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', Axis::TICK_LABEL_POSITION_NEXT_TO);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisY()->setTickLabelPosition(Axis::TICK_LABEL_POSITION_HIGH);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', Axis::TICK_LABEL_POSITION_HIGH);
        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisY()->setTickLabelPosition(Axis::TICK_LABEL_POSITION_LOW);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', Axis::TICK_LABEL_POSITION_LOW);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisTickMark(): void
    {
        $value = Axis::TICK_MARK_CROSS;

        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFill()->setStartColor(new Color('FFAABBCC'));
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $elementMax = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:majorTickMark';
        $elementMin = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:minorTickMark';

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', Axis::TICK_MARK_NONE);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin, 'val', Axis::TICK_MARK_NONE);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisY()->setMinorTickMark($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', Axis::TICK_MARK_NONE);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin, 'val', $value);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisY()->setMinorTickMark();
        $oShape->getPlotArea()->getAxisY()->setMajorTickMark($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin, 'val', Axis::TICK_MARK_NONE);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', $value);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisY()->setMinorTickMark($value);
        $oShape->getPlotArea()->getAxisY()->setMajorTickMark($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin, 'val', $value);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', $value);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisTitleRotation(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);

        $oBar = new Bar();
        $oBar->addSeries($oSeries);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->getPlotArea()->setType($oBar);

        $pathShape = 'ppt/charts/' . $oShape->getIndexedFilename();
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:bodyPr';
        $this->assertZipXmlElementExists($pathShape, $element);
        $this->assertZipXmlAttributeNotExists($pathShape, $element, 'rot');

        $this->resetPresentationFile();
        $value = rand(0, 360);
        $oShape->getPlotArea()->getAxisX()->setTitleRotation($value);

        $pathShape = 'ppt/charts/' . $oShape->getIndexedFilename();
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:title/c:tx/c:rich/a:bodyPr';
        $this->assertZipXmlElementExists($pathShape, $element);
        $this->assertZipXmlAttributeExists($pathShape, $element, 'rot');
        $this->assertZipXmlAttributeEquals($pathShape, $element, 'rot', Drawing::degreesToAngle((int) $value));
    }

    public function testAxisVisibility(): void
    {
        $element = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:delete';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oLine = new Line();
        $oShape->getPlotArea()->setType($oLine);

        // Set Visible : FALSE
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Axis', $oShape->getPlotArea()->getAxisX()->setIsVisible(false));
        $this->assertFalse($oShape->getPlotArea()->getAxisX()->isVisible());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '1');

        $this->assertIsSchemaECMA376Valid();

        $this->resetPresentationFile();

        // Set Visible : TRUE
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Axis', $oShape->getPlotArea()->getAxisX()->setIsVisible(true));
        $this->assertTrue($oShape->getPlotArea()->getAxisX()->isVisible());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '0');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisXUnit(): void
    {
        $value = mt_rand(0, 100);

        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $elementMax = '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:tickLblSkip';

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisX()->setMajorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', $value);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testAxisYUnit(): void
    {
        $value = mt_rand(0, 100);

        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $elementMax = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:majorUnit';
        $elementMin = '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:minorUnit';

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisY()->setMinorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin, 'val', $value);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisY()->setMinorUnit(null);
        $oShape->getPlotArea()->getAxisY()->setMajorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', $value);

        $this->assertIsSchemaECMA376Valid();

        $oShape->getPlotArea()->getAxisY()->setMinorUnit($value);
        $oShape->getPlotArea()->getAxisY()->setMajorUnit($value);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMin, 'val', $value);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $elementMax, 'val', $value);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeArea(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oArea = new Area();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFill()->setStartColor(new Color('FFAABBCC'));
        $oArea->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oArea);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:areaChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:areaChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeBar(): void
    {
        $valueGapWidthPercent = mt_rand(0, 500);
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar = new Bar();
        $oBar->setGapWidthPercent($valueGapWidthPercent);
        $oBar->setOverlapWidthPercent(-10);
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oBar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser/c:dPt/c:spPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $oSeries->getTitle());
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart/c:gapWidth';
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', $valueGapWidthPercent);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:barChart/c:overlap';
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', $oBar->getOverlapWidthPercent());

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeBar3D(): void
    {
        $valueGapWidthPercent = mt_rand(0, 500);
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oBar3D->setGapWidthPercent($valueGapWidthPercent);
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:dPt/c:spPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $oSeries->getTitle());
        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:gapWidth';
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', $valueGapWidthPercent);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeBar3DSubScript(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'baseline', '-250000');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeBar3DSuperScript(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'baseline', '300000');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeDoughnut(): void
    {
        $randHoleSize = mt_rand(10, 90);
        $randSeparator = chr(rand(ord('A'), ord('Z')));

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

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:ser/c:dPt/c:spPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $oSeries->getTitle());
        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:dLbls/c:separator';
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:dLbls/c:showBubbleSize';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 0);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:firstSliceAng';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 0);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:holeSize';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 50);

        $oDoughnut->setHoleSize($randHoleSize);
        $this->resetPresentationFile();

        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:holeSize';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', $randHoleSize);

        $oSeries->setSeparator($randSeparator);
        $this->resetPresentationFile();

        $element = '/c:chartSpace/c:chart/c:plotArea/c:doughnutChart/c:dLbls/c:separator';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $randSeparator);
    }

    public function testTypeBar3DBarDirection(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oBar3D = new Bar3D();
        $oBar3D->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);
        $oSeries = new Series('Downloads', $this->seriesData);
        $oBar3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oBar3D);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:barDir';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', Bar3D::DIRECTION_HORIZONTAL);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeLine(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $oSeries->getTitle());

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeLineGridlines(): void
    {
        $arrayTests = [
            [
                'methodAxis' => 'getAxisX',
                'methodGrid' => 'setMajorGridlines',
                'expectedElement' => '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:majorGridlines/c:spPr/a:ln',
                'expectedElementColor' => '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:majorGridlines/c:spPr/a:ln/a:solidFill/a:srgbClr',
            ],
            [
                'methodAxis' => 'getAxisX',
                'methodGrid' => 'setMinorGridlines',
                'expectedElement' => '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:minorGridlines/c:spPr/a:ln',
                'expectedElementColor' => '/c:chartSpace/c:chart/c:plotArea/c:catAx/c:minorGridlines/c:spPr/a:ln/a:solidFill/a:srgbClr',
            ],
            [
                'methodAxis' => 'getAxisY',
                'methodGrid' => 'setMajorGridlines',
                'expectedElement' => '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:majorGridlines/c:spPr/a:ln',
                'expectedElementColor' => '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:majorGridlines/c:spPr/a:ln/a:solidFill/a:srgbClr',
            ],
            [
                'methodAxis' => 'getAxisY',
                'methodGrid' => 'setMinorGridlines',
                'expectedElement' => '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:minorGridlines/c:spPr/a:ln',
                'expectedElementColor' => '/c:chartSpace/c:chart/c:plotArea/c:valAx/c:minorGridlines/c:spPr/a:ln/a:solidFill/a:srgbClr',
            ],
        ];
        $expectedColor = new Color(Color::COLOR_BLUE);
        foreach ($arrayTests as $arrayTest) {
            $expectedSizePts = mt_rand(1, 100);
            $expectedSizeEmu = Drawing::pointsToEmu($expectedSizePts);

            $this->oPresentation->removeSlideByIndex()->createSlide();
            $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
            $oLine = new Line();
            $oLine->addSeries(new Series('Downloads', [
                'A' => '1',
                'B' => '2',
                'C' => '4',
                'D' => '3',
                'E' => '2',
            ]));
            $oShape->getPlotArea()->setType($oLine);
            $oGridlines = new Gridlines();
            $oGridlines->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($expectedColor);
            $oGridlines->getOutline()->setWidth($expectedSizePts);
            $oShape->getPlotArea()->{$arrayTest['methodAxis']}()->{$arrayTest['methodGrid']}($oGridlines);

            $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
            $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $arrayTest['expectedElement']);
            $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $arrayTest['expectedElement'], 'w');
            $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $arrayTest['expectedElement'], 'w', $expectedSizeEmu);
            $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $arrayTest['expectedElementColor']);
            $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $arrayTest['expectedElementColor'], 'val');
            $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $arrayTest['expectedElementColor'], 'val', $expectedColor->getRGB());

            $this->assertIsSchemaECMA376Valid();

            $this->resetPresentationFile();
        }
    }

    /**
     * @dataProvider dataProviderMarkerSymbol
     */
    public function testTypeLineMarker(string $expectedSymbol): void
    {
        $expectedSize = mt_rand(2, 72);
        $expectedEltSymbol = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:marker/c:symbol';
        $expectedElementSize = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:marker/c:size';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getMarker()->setSymbol($expectedSymbol)->setSize($expectedSize);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedEltSymbol);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedEltSymbol, 'val', $expectedSymbol);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize, 'val', $expectedSize);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->getMarker()->setSize(1);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize, 'val', 2);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->getMarker()->setSize(73);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize, 'val', 72);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->getMarker()->setSymbol(Marker::SYMBOL_NONE);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedEltSymbol);
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeLineMarkerBorder(): void
    {
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:marker/c:spPr/a:ln';

        $oSeries = new Series('Downloads', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'w', 12700);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'cmpd', Border::LINE_SINGLE);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element . '/a:prstDash');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element . '/a:prstDash', 'val', Border::DASH_SOLID);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeLineMarkerFill(): void
    {
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:marker/c:spPr/a:noFill';

        $oSeries = new Series('Downloads', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeLineSeriesOutline(): void
    {
        $expectedWidth = mt_rand(1, 100);
        $expectedWidthEmu = Drawing::pointsToEmu($expectedWidth);
        $expectedElement = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:spPr/a:ln';

        $oOutline = new Outline();
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_YELLOW));
        $oOutline->setWidth($expectedWidth);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->setOutline($oOutline);
        $oLine->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement, 'w', $expectedWidthEmu);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement . '/a:solidFill');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeLineSmooth(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);

        $oLine = new Line();
        $oLine->addSeries($oSeries);
        $oLine->setIsSmooth(false);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oShape->getPlotArea()->setType($oLine);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:smooth';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '0');

        $this->assertIsSchemaECMA376Valid();

        $this->resetPresentationFile();
        $oLine->setIsSmooth(true);
        $oShape->getPlotArea()->setType($oLine);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:smooth';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '1');
    }

    public function testTypeLineSubScript(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'baseline', '-250000');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeLineSuperScript(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oLine = new Line();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oLine->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oLine);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:lineChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'baseline', '300000');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypePie(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie = new Pie();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oPie->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser/c:dPt/c:spPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $oSeries->getTitle());
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser/c:dLbls/c:showLegendKey';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 0);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->setShowLegendKey(true);
        $this->resetPresentationFile();

        $element = '/c:chartSpace/c:chart/c:plotArea/c:pieChart/c:ser/c:dLbls/c:showLegendKey';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 1);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypePie3D(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKBLUE));
        $oSeries->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));
        $oSeries->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKRED));
        $oSeries->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKYELLOW));
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:dPt/c:spPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $oSeries->getTitle());

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypePie3DExplosion(): void
    {
        $value = mt_rand(1, 100);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oPie3D->setExplosion($value);
        $oSeries = new Series('Downloads', $this->seriesData);
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:explosion';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', $value);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypePie3DSubScript(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'baseline', '-250000');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypePie3DSuperScript(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oPie3D = new Pie3D();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oPie3D->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oPie3D);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'baseline', '300000');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeRadar(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oRadar = new Radar();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oRadar->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oRadar);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:radarChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:radarChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:radarChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $oSeries->getTitle());
        $element = '/c:chartSpace/c:chart/c:plotArea/c:radarChart/c:ser/c:dLbls/c:showLegendKey';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 0);
    }

    public function testTypeScatter(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $element = '/p:sld/p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData';
        $this->assertZipXmlElementExists('ppt/slides/slide1.xml', $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:tx/c:v';
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, $oSeries->getTitle());
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:showLegendKey';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 0);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->setShowLegendKey(true);
        $this->resetPresentationFile();

        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:showLegendKey';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', 1);

        $this->assertIsSchemaECMA376Valid();
    }

    /**
     * @dataProvider dataProviderMarkerSymbol
     */
    public function testTypeScatterMarkerSymbol(string $expectedSymbol): void
    {
        $expectedSize = mt_rand(2, 72);
        $expectedEltSymbol = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:marker/c:symbol';
        $expectedElementSize = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:marker/c:size';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getMarker()->setSymbol($expectedSymbol)->setSize($expectedSize);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedEltSymbol);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedEltSymbol, 'val', $expectedSymbol);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize, 'val', $expectedSize);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->getMarker()->setSize(1);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize, 'val', 2);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->getMarker()->setSize(73);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize, 'val', 72);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->getMarker()->setSymbol(Marker::SYMBOL_NONE);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedEltSymbol);
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElementSize);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeScatterMarkerBorder(): void
    {
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:marker/c:spPr/a:ln';

        $oSeries = new Series('Downloads', $this->seriesData);

        $oScatter = new Scatter();
        $oScatter->addSeries($oSeries);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->getPlotArea()->setType($oScatter);

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'w', 12700);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'cmpd', Border::LINE_SINGLE);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element . '/a:prstDash');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element . '/a:prstDash', 'val', Border::DASH_SOLID);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeScatterMarkerFill(): void
    {
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:marker/c:spPr/a:noFill';

        $oSeries = new Series('Downloads', $this->seriesData);

        $oScatter = new Scatter();
        $oScatter->addSeries($oSeries);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->getPlotArea()->setType($oScatter);

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeScatterSeparator(): void
    {
        $expectedSeparator = ';';
        $expectedElement = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:separator';

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement);
        $this->assertIsSchemaECMA376Valid();

        $oSeries->setSeparator($expectedSeparator);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement);
        $this->assertZipXmlElementEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement, $expectedSeparator);
        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeScatterSeriesOutline(): void
    {
        $expectedWidth = mt_rand(1, 100);
        $expectedWidthEmu = Drawing::pointsToEmu($expectedWidth);
        $expectedElement = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:spPr/a:ln';

        $oOutline = new Outline();
        // Define the color
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_YELLOW));
        // Define the width (in points)
        $oOutline->setWidth($expectedWidth);

        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement);

        $this->assertIsSchemaECMA376Valid();

        $oSeries->setOutline($oOutline);
        $oScatter->setSeries([$oSeries]);
        $this->resetPresentationFile();

        $this->assertZipFileExists('ppt/charts/' . $oShape->getIndexedFilename());
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement, 'w', $expectedWidthEmu);
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $expectedElement . '/a:solidFill');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeScatterSmooth(): void
    {
        $oSeries = new Series('Downloads', $this->seriesData);

        $oScatter = new Scatter();
        $oScatter->addSeries($oSeries);
        $oScatter->setIsSmooth(false);

        $oShape = $this->oPresentation->getActiveSlide()->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oShape->getPlotArea()->setType($oScatter);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:smooth';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '0');

        $this->assertIsSchemaECMA376Valid();

        $this->resetPresentationFile();
        $oScatter->setIsSmooth(true);
        $oShape->getPlotArea()->setType($oScatter);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:smooth';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeExists('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val');
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '1');
    }

    public function testTypeScatterSubScript(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSubScript(true);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'baseline', '-250000');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testTypeScatterSuperScript(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createChartShape();
        $oShape->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $oScatter = new Scatter();
        $oSeries = new Series('Downloads', $this->seriesData);
        $oSeries->getFont()->setSuperScript(true);
        $oScatter->addSeries($oSeries);
        $oShape->getPlotArea()->setType($oScatter);

        $element = '/c:chartSpace/c:chart/c:plotArea/c:scatterChart/c:ser/c:dLbls/c:txPr/a:p/a:pPr/a:defRPr';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'baseline', '300000');

        $this->assertIsSchemaECMA376Valid();
    }

    public function testView3D(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oLine = new Line();
        $oLine->addSeries(new Series('Downloads', $this->seriesData));
        $oShape = $oSlide->createChartShape();
        $oShape->getPlotArea()->setType($oLine);

        $element = '/c:chartSpace/c:chart/c:view3D/c:hPercent';
        $this->assertZipXmlElementExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertZipXmlAttributeEquals('ppt/charts/' . $oShape->getIndexedFilename(), $element, 'val', '100');
        $this->assertIsSchemaECMA376Valid();

        $oShape->getView3D()->setHeightPercent(null);
        $this->resetPresentationFile();

        $this->assertZipXmlElementNotExists('ppt/charts/' . $oShape->getIndexedFilename(), $element);
        $this->assertIsSchemaECMA376Valid();
    }

    /**
     * @return array<array<string>>
     */
    public function dataProviderMarkerSymbol(): iterable
    {
        foreach (Marker::$arraySymbol as $symbol) {
            if ($symbol === Marker::SYMBOL_NONE) {
                continue;
            }
            yield [$symbol];
        }
    }
}
