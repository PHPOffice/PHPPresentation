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
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Pie element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Type\Pie
 */
class PieTest extends TestCase
{
    public function testData(): void
    {
        $object = new Pie();

        self::assertIsArray($object->getSeries());
        self::assertEmpty($object->getSeries());

        $array = [
            new Series(),
            new Series(),
        ];

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie', $object->setSeries());
        self::assertEmpty($object->getSeries());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie', $object->setSeries($array));
        self::assertCount(count($array), $object->getSeries());
    }

    public function testSeries(): void
    {
        $object = new Pie();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie', $object->addSeries(new Series()));
        self::assertCount(1, $object->getSeries());
    }

    public function testHashCode(): void
    {
        $oSeries = new Series();

        $object = new Pie();
        $object->addSeries($oSeries);

        self::assertEquals(md5($oSeries->getHashCode() . get_class($object)), $object->getHashCode());
    }
}
