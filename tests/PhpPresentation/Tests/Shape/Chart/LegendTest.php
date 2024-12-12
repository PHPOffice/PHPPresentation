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

use PhpOffice\PhpPresentation\Shape\Chart\Legend;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Legend element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Legend
 */
class LegendTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Legend();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getBorder());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->getAlignment());
    }

    public function testAlignment(): void
    {
        $object = new Legend();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setAlignment(new Alignment()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->getAlignment());
    }

    public function testBorder(): void
    {
        $object = new Legend();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getBorder());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setBorder(new Border()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getBorder());
    }

    public function testFill(): void
    {
        $object = new Legend();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setFill(new Fill()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
    }

    public function testFont(): void
    {
        $object = new Legend();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setFont());
        self::assertNull($object->getFont());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setFont(new Font()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
    }

    public function testHashIndex(): void
    {
        $object = new Legend();
        $value = mt_rand(1, 100);

        self::assertEmpty($object->getHashIndex());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setHashIndex($value));
        self::assertEquals($value, $object->getHashIndex());
    }

    public function testHeight(): void
    {
        $object = new Legend();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setHeight());
        self::assertEquals(0, $object->getHeight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setHeight($value));
        self::assertEquals($value, $object->getHeight());
    }

    public function testOffsetX(): void
    {
        $object = new Legend();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setOffsetX());
        self::assertEquals(0, $object->getOffsetX());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setOffsetX($value));
        self::assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY(): void
    {
        $object = new Legend();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setOffsetY());
        self::assertEquals(0, $object->getOffsetY());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setOffsetY($value));
        self::assertEquals($value, $object->getOffsetY());
    }

    public function testPosition(): void
    {
        $object = new Legend();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setPosition());
        self::assertEquals(Legend::POSITION_RIGHT, $object->getPosition());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setPosition(Legend::POSITION_BOTTOM));
        self::assertEquals(Legend::POSITION_BOTTOM, $object->getPosition());
    }

    public function testVisible(): void
    {
        $object = new Legend();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setVisible());
        self::assertTrue($object->isVisible());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setVisible(true));
        self::assertTrue($object->isVisible());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setVisible(false));
        self::assertFalse($object->isVisible());
    }

    public function testWidth(): void
    {
        $object = new Legend();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setWidth());
        self::assertEquals(0, $object->getWidth());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Legend', $object->setWidth($value));
        self::assertEquals($value, $object->getWidth());
    }
}
