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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart\Type;

use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Doughnut;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Doughnut element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Type\Doughnut
 */
class DoughnutTest extends TestCase
{
    public function testData(): void
    {
        $object = new Doughnut();

        self::assertIsArray($object->getSeries());
        self::assertEmpty($object->getSeries());

        $array = [
            new Series(),
            new Series(),
        ];

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Doughnut', $object->setSeries());
        self::assertEmpty($object->getSeries());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Doughnut', $object->setSeries($array));
        self::assertCount(count($array), $object->getSeries());
    }

    public function testHoleSize(): void
    {
        $rand = mt_rand(10, 90);
        $object = new Doughnut();

        self::assertEquals(50, $object->getHoleSize());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Doughnut', $object->setHoleSize(9));
        self::assertEquals(10, $object->getHoleSize());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Doughnut', $object->setHoleSize(91));
        self::assertEquals(90, $object->getHoleSize());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Doughnut', $object->setHoleSize($rand));
        self::assertEquals($rand, $object->getHoleSize());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Doughnut', $object->setHoleSize());
        self::assertEquals(50, $object->getHoleSize());
    }

    public function testSeries(): void
    {
        $object = new Doughnut();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Doughnut', $object->addSeries(new Series()));
        self::assertCount(1, $object->getSeries());
    }

    public function testHashCode(): void
    {
        $oSeries = new Series();

        $object = new Doughnut();
        $object->addSeries($oSeries);

        self::assertEquals(md5($oSeries->getHashCode() . get_class($object)), $object->getHashCode());
    }

    public function testFirstSliceAngle(): void
    {
        $doughnut = new Doughnut();

        // 1) default
        self::assertSame(0, $doughnut->getFirstSliceAngle());

        // 2) fluent + simple set/get
        $angle = $doughnut->setFirstSliceAngle(90);
        self::assertInstanceOf(Doughnut::class, $angle);
        self::assertSame(90, $doughnut->getFirstSliceAngle());

        // 3) normalization (overflow wraps)
        $doughnut->setFirstSliceAngle(450); // 450 % 360 = 90
        self::assertSame(90, $doughnut->getFirstSliceAngle());

        // 3) normalization (negative wraps)
        $doughnut->setFirstSliceAngle(-45); // -> 315
        self::assertSame(315, $doughnut->getFirstSliceAngle());
    }
}
