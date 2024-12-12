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

        self::assertInstanceOf(Fill::class, $object->getFill());
        self::assertInstanceOf(Font::class, $object->getFont());
        self::assertEquals('Calibri', $object->getFont()->getName());
        self::assertEquals(9, $object->getFont()->getSize());
        self::assertEquals('Series Title', $object->getTitle());
        self::assertIsArray($object->getValues());
        self::assertEmpty($object->getValues());
        self::assertInstanceOf(Marker::class, $object->getMarker());
        self::assertNull($object->getOutline());
        self::assertFalse($object->hasShowLegendKey());
    }

    public function testDataLabelNumFormat(): void
    {
        $object = new Series();

        self::assertEmpty($object->getDlblNumFormat());
        self::assertFalse($object->hasDlblNumFormat());

        self::assertInstanceOf(Series::class, $object->setDlblNumFormat('#%'));

        self::assertEquals('#%', $object->getDlblNumFormat());
        self::assertTrue($object->hasDlblNumFormat());

        self::assertInstanceOf(Series::class, $object->setDlblNumFormat());

        self::assertEmpty($object->getDlblNumFormat());
        self::assertFalse($object->hasDlblNumFormat());
    }

    public function testDataPointFills(): void
    {
        $object = new Series();

        self::assertIsArray($object->getDataPointFills());
        self::assertEmpty($object->getDataPointFills());

        self::assertInstanceOf(Fill::class, $object->getDataPointFill(0));
    }

    public function testFill(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setFill());
        self::assertNull($object->getFill());
        self::assertInstanceOf(Series::class, $object->setFill(new Fill()));
        self::assertInstanceOf(Fill::class, $object->getFill());
    }

    public function testFont(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setFont());
        self::assertNull($object->getFont());
        self::assertInstanceOf(Series::class, $object->setFont(new Font()));
        self::assertInstanceOf(Font::class, $object->getFont());
    }

    public function testHashIndex(): void
    {
        $object = new Series();
        $value = mt_rand(1, 100);

        self::assertEmpty($object->getHashIndex());
        self::assertInstanceOf(Series::class, $object->setHashIndex($value));
        self::assertEquals($value, $object->getHashIndex());
    }

    public function testHashCode(): void
    {
        $object = new Series();

        self::assertEquals(md5($object->getFill()->getHashCode() . $object->getFont()->getHashCode() . var_export($object->getValues(), true) . var_export($object, true) . get_class($object)), $object->getHashCode());
    }

    public function testLabelPosition(): void
    {
        $object = new Series();

        self::assertEmpty($object->getHashIndex());
        self::assertInstanceOf(Series::class, $object->setLabelPosition(Series::LABEL_INSIDEBASE));
        self::assertEquals(Series::LABEL_INSIDEBASE, $object->getLabelPosition());
    }

    public function testMarker(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setMarker(new Marker()));
        self::assertInstanceOf(Marker::class, $object->getMarker());
    }

    public function testOutline(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setOutline(new Outline()));
        self::assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testShowCategoryName(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setShowCategoryName(true));
        self::assertTrue($object->hasShowCategoryName());
        self::assertInstanceOf(Series::class, $object->setShowCategoryName(false));
        self::assertFalse($object->hasShowCategoryName());
    }

    public function testShowLeaderLines(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setShowLeaderLines(true));
        self::assertTrue($object->hasShowLeaderLines());
        self::assertInstanceOf(Series::class, $object->setShowLeaderLines(false));
        self::assertFalse($object->hasShowLeaderLines());
    }

    public function testShowLegendKey(): void
    {
        $object = new Series();

        self::assertFalse($object->hasShowLegendKey());
        self::assertInstanceOf(Series::class, $object->setShowLegendKey(true));
        self::assertTrue($object->hasShowLegendKey());
        self::assertInstanceOf(Series::class, $object->setShowLegendKey(false));
        self::assertFalse($object->hasShowLegendKey());
    }

    public function testShowPercentage(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setShowPercentage(true));
        self::assertTrue($object->hasShowPercentage());
        self::assertInstanceOf(Series::class, $object->setShowPercentage(false));
        self::assertFalse($object->hasShowPercentage());
    }

    public function testShowSeparator(): void
    {
        $value = ';';
        $object = new Series();

        self::assertFalse($object->hasShowSeparator());
        self::assertInstanceOf(Series::class, $object->setSeparator($value));
        self::assertEquals($value, $object->getSeparator());
        self::assertTrue($object->hasShowSeparator());
        self::assertInstanceOf(Series::class, $object->setSeparator(''));
        self::assertFalse($object->hasShowPercentage());
    }

    public function testShowSeriesName(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setShowSeriesName(true));
        self::assertTrue($object->hasShowSeriesName());
        self::assertInstanceOf(Series::class, $object->setShowSeriesName(false));
        self::assertFalse($object->hasShowSeriesName());
    }

    public function testShowValue(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setShowValue(true));
        self::assertTrue($object->hasShowValue());
        self::assertInstanceOf(Series::class, $object->setShowValue(false));
        self::assertFalse($object->hasShowValue());
    }

    public function testTitle(): void
    {
        $object = new Series();

        self::assertInstanceOf(Series::class, $object->setTitle());
        self::assertEquals('Series Title', $object->getTitle());
        self::assertInstanceOf(Series::class, $object->setTitle('AAAA'));
        self::assertEquals('AAAA', $object->getTitle());
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

        self::assertIsArray($object->getValues());
        self::assertEmpty($object->getValues());
        self::assertInstanceOf(Series::class, $object->setValues());
        self::assertEmpty($object->getValues());
        self::assertInstanceOf(Series::class, $object->setValues($array));
        self::assertCount(count($array), $object->getValues());
        self::assertInstanceOf(Series::class, $object->addValue('4', 'e'));
        self::assertCount(count($array) + 1, $object->getValues());
    }

    public function testClone(): void
    {
        $object = new Series();
        $object->setOutline(new Outline());
        $clone = clone $object;

        self::assertInstanceOf(Series::class, $clone);
        self::assertInstanceOf(Outline::class, $clone->getOutline());
    }
}
