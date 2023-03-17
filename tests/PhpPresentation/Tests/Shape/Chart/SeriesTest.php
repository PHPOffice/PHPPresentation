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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Series element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Series
 */
class SeriesTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Fill::class, $object->getFill());
        $this->assertInstanceOf(Font::class, $object->getFont());
        $this->assertEquals('Calibri', $object->getFont()->getName());
        $this->assertEquals(9, $object->getFont()->getSize());
        $this->assertEquals('Series Title', $object->getTitle());
        $this->assertIsArray($object->getValues());
        $this->assertEmpty($object->getValues());
        $this->assertInstanceOf(Marker::class, $object->getMarker());
        $this->assertNull($object->getOutline());
        $this->assertFalse($object->hasShowLegendKey());
    }

    public function testDataLabelNumFormat(): void
    {
        $object = new Series();

        $this->assertEmpty($object->getDlblNumFormat());
        $this->assertFalse($object->hasDlblNumFormat());

        $this->assertInstanceOf(Series::class, $object->setDlblNumFormat('#%'));

        $this->assertEquals('#%', $object->getDlblNumFormat());
        $this->assertTrue($object->hasDlblNumFormat());

        $this->assertInstanceOf(Series::class, $object->setDlblNumFormat());

        $this->assertEmpty($object->getDlblNumFormat());
        $this->assertFalse($object->hasDlblNumFormat());
    }

    public function testDataPointFills(): void
    {
        $object = new Series();

        $this->assertIsArray($object->getDataPointFills());
        $this->assertEmpty($object->getDataPointFills());

        $this->assertInstanceOf(Fill::class, $object->getDataPointFill(0));
    }

    public function testFill(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setFill());
        $this->assertNull($object->getFill());
        $this->assertInstanceOf(Series::class, $object->setFill(new Fill()));
        $this->assertInstanceOf(Fill::class, $object->getFill());
    }

    public function testFont(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setFont());
        $this->assertNull($object->getFont());
        $this->assertInstanceOf(Series::class, $object->setFont(new Font()));
        $this->assertInstanceOf(Font::class, $object->getFont());
    }

    public function testHashIndex(): void
    {
        $object = new Series();
        $value = mt_rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf(Series::class, $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testHashCode(): void
    {
        $object = new Series();

        $this->assertEquals(md5($object->getFill()->getHashCode() . $object->getFont()->getHashCode() . var_export($object->getValues(), true) . var_export($object, true) . get_class($object)), $object->getHashCode());
    }

    public function testLabelPosition(): void
    {
        $object = new Series();

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf(Series::class, $object->setLabelPosition(Series::LABEL_INSIDEBASE));
        $this->assertEquals(Series::LABEL_INSIDEBASE, $object->getLabelPosition());
    }

    public function testMarker(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setMarker(new Marker()));
        $this->assertInstanceOf(Marker::class, $object->getMarker());
    }

    public function testOutline(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setOutline(new Outline()));
        $this->assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testShowCategoryName(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setShowCategoryName(true));
        $this->assertTrue($object->hasShowCategoryName());
        $this->assertInstanceOf(Series::class, $object->setShowCategoryName(false));
        $this->assertFalse($object->hasShowCategoryName());
    }

    public function testShowLeaderLines(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setShowLeaderLines(true));
        $this->assertTrue($object->hasShowLeaderLines());
        $this->assertInstanceOf(Series::class, $object->setShowLeaderLines(false));
        $this->assertFalse($object->hasShowLeaderLines());
    }

    public function testShowLegendKey(): void
    {
        $object = new Series();

        $this->assertFalse($object->hasShowLegendKey());
        $this->assertInstanceOf(Series::class, $object->setShowLegendKey(true));
        $this->assertTrue($object->hasShowLegendKey());
        $this->assertInstanceOf(Series::class, $object->setShowLegendKey(false));
        $this->assertFalse($object->hasShowLegendKey());
    }

    public function testShowPercentage(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setShowPercentage(true));
        $this->assertTrue($object->hasShowPercentage());
        $this->assertInstanceOf(Series::class, $object->setShowPercentage(false));
        $this->assertFalse($object->hasShowPercentage());
    }

    public function testShowSeparator(): void
    {
        $value = ';';
        $object = new Series();

        $this->assertFalse($object->hasShowSeparator());
        $this->assertInstanceOf(Series::class, $object->setSeparator($value));
        $this->assertEquals($value, $object->getSeparator());
        $this->assertTrue($object->hasShowSeparator());
        $this->assertInstanceOf(Series::class, $object->setSeparator(''));
        $this->assertFalse($object->hasShowPercentage());
    }

    public function testShowSeriesName(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setShowSeriesName(true));
        $this->assertTrue($object->hasShowSeriesName());
        $this->assertInstanceOf(Series::class, $object->setShowSeriesName(false));
        $this->assertFalse($object->hasShowSeriesName());
    }

    public function testShowValue(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setShowValue(true));
        $this->assertTrue($object->hasShowValue());
        $this->assertInstanceOf(Series::class, $object->setShowValue(false));
        $this->assertFalse($object->hasShowValue());
    }

    public function testTitle(): void
    {
        $object = new Series();

        $this->assertInstanceOf(Series::class, $object->setTitle());
        $this->assertEquals('Series Title', $object->getTitle());
        $this->assertInstanceOf(Series::class, $object->setTitle('AAAA'));
        $this->assertEquals('AAAA', $object->getTitle());
    }

    public function testValue(): void
    {
        $object = new Series();

        /** @var array<string, string> $array */
        $array = [
            '0' => 'a',
            '1' => 'b',
            '2' => 'c',
            '3' => 'd',
        ];

        $this->assertIsArray($object->getValues());
        $this->assertEmpty($object->getValues());
        $this->assertInstanceOf(Series::class, $object->setValues());
        $this->assertEmpty($object->getValues());
        $this->assertInstanceOf(Series::class, $object->setValues($array));
        $this->assertCount(count($array), $object->getValues());
        $this->assertInstanceOf(Series::class, $object->addValue('4', 'e'));
        $this->assertCount(count($array) + 1, $object->getValues());
    }

    public function testClone(): void
    {
        $object = new Series();
        $object->setOutline(new Outline());
        $clone = clone $object;

        $this->assertInstanceOf(Series::class, $clone);
        $this->assertInstanceOf(Outline::class, $clone->getOutline());
    }
}
