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
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Legend element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Marker
 */
class MarkerTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Marker();

        self::assertEquals(Marker::SYMBOL_NONE, $object->getSymbol());
        self::assertEquals(5, $object->getSize());
    }

    public function testBorder(): void
    {
        $object = new Marker();

        self::assertInstanceOf(Border::class, $object->getBorder());
        self::assertInstanceOf(Marker::class, $object->setBorder(new Border()));
    }

    public function testFill(): void
    {
        $object = new Marker();

        self::assertInstanceOf(Fill::class, $object->getFill());
        self::assertInstanceOf(Marker::class, $object->setFill(new Fill()));
    }

    public function testSize(): void
    {
        $object = new Marker();
        $value = mt_rand(1, 100);

        self::assertInstanceOf(Marker::class, $object->setSize($value));
        self::assertEquals($value, $object->getSize());
    }

    public function testSymbol(): void
    {
        $object = new Marker();

        $expected = Marker::$arraySymbol[array_rand(Marker::$arraySymbol)];

        self::assertInstanceOf(Marker::class, $object->setSymbol($expected));
        self::assertEquals($expected, $object->getSymbol());
    }
}
