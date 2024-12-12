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

use PhpOffice\PhpPresentation\Shape\Chart\Title;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Font;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Title element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Title
 */
class TitleTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Title();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->getAlignment());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
        self::assertEquals('Calibri', $object->getFont()->getName());
        self::assertEquals(18, $object->getFont()->getSize());
    }

    public function testAlignment(): void
    {
        $object = new Title();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setAlignment(new Alignment()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Alignment', $object->getAlignment());
    }

    public function testFont(): void
    {
        $object = new Title();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setFont());
        self::assertNull($object->getFont());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setFont(new Font()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
    }

    public function testHashIndex(): void
    {
        $object = new Title();
        $value = mt_rand(1, 100);

        self::assertEmpty($object->getHashIndex());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setHashIndex($value));
        self::assertEquals($value, $object->getHashIndex());
    }

    public function testHeight(): void
    {
        $object = new Title();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setHeight());
        self::assertEquals(0, $object->getHeight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setHeight($value));
        self::assertEquals($value, $object->getHeight());
    }

    public function testOffsetX(): void
    {
        $object = new Title();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setOffsetX());
        self::assertEquals(0.01, $object->getOffsetX());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setOffsetX($value));
        self::assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY(): void
    {
        $object = new Title();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setOffsetY());
        self::assertEquals(0.01, $object->getOffsetY());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setOffsetY($value));
        self::assertEquals($value, $object->getOffsetY());
    }

    public function testText(): void
    {
        $object = new Title();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setText());
        self::assertNull($object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setText('AAAA'));
        self::assertEquals('AAAA', $object->getText());
    }

    public function testVisible(): void
    {
        $object = new Title();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setVisible());
        self::assertTrue($object->isVisible());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setVisible(true));
        self::assertTrue($object->isVisible());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setVisible(false));
        self::assertFalse($object->isVisible());
    }

    public function testWidth(): void
    {
        $object = new Title();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setWidth());
        self::assertEquals(0, $object->getWidth());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Title', $object->setWidth($value));
        self::assertEquals($value, $object->getWidth());
    }
}
