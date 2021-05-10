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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class BorderTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Border();
        static::assertEquals(1, $object->getLineWidth());
        static::assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        static::assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        static::assertEquals('FF000000', $object->getColor()->getARGB());
    }

    /**
     * Test get/set color
     */
    public function testSetGetColor()
    {
        $object = new Border();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setColor());
        static::assertNull($object->getColor());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setColor(new Color(Color::COLOR_BLUE)));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        static::assertEquals('FF0000FF', $object->getColor()->getARGB());
    }

    /**
     * Test get/set dash style
     */
    public function testSetGetDashStyle()
    {
        $object = new Border();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle());
        static::assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle(''));
        static::assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle(Border::DASH_DASH));
        static::assertEquals(Border::DASH_DASH, $object->getDashStyle());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Border();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        static::assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set line style
     */
    public function testSetGetLineStyle()
    {
        $object = new Border();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle());
        static::assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle(''));
        static::assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle(Border::LINE_DOUBLE));
        static::assertEquals(Border::LINE_DOUBLE, $object->getLineStyle());
    }

    /**
     * Test get/set line width
     */
    public function testSetGetLineWidth()
    {
        $object = new Border();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineWidth());
        static::assertEquals(1, $object->getLineWidth());
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineWidth($value));
        static::assertEquals($value, $object->getLineWidth());
    }
}
