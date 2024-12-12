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

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class BorderTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $object = new Border();
        self::assertEquals(1, $object->getLineWidth());
        self::assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        self::assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        self::assertEquals('FF000000', $object->getColor()->getARGB());
    }

    /**
     * Test get/set color.
     */
    public function testSetGetColor(): void
    {
        $object = new Border();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setColor());
        self::assertNull($object->getColor());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setColor(new Color(Color::COLOR_BLUE)));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        self::assertEquals('FF0000FF', $object->getColor()->getARGB());
    }

    /**
     * Test get/set dash style.
     */
    public function testSetGetDashStyle(): void
    {
        $object = new Border();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle());
        self::assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle(''));
        self::assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle(Border::DASH_DASH));
        self::assertEquals(Border::DASH_DASH, $object->getDashStyle());
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $object = new Border();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set line style.
     */
    public function testSetGetLineStyle(): void
    {
        $object = new Border();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle());
        self::assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle(''));
        self::assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle(Border::LINE_DOUBLE));
        self::assertEquals(Border::LINE_DOUBLE, $object->getLineStyle());
    }

    /**
     * Test get/set line width.
     */
    public function testSetGetLineWidth(): void
    {
        $object = new Border();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineWidth());
        self::assertEquals(1, $object->getLineWidth());
        $value = mt_rand(1, 100);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineWidth($value));
        self::assertEquals($value, $object->getLineWidth());
    }
}
