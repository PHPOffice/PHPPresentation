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
        $this->assertEquals(1, $object->getLineWidth());
        $this->assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        $this->assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertEquals('FF000000', $object->getColor()->getARGB());
    }

    /**
     * Test get/set color.
     */
    public function testSetGetColor(): void
    {
        $object = new Border();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setColor());
        $this->assertNull($object->getColor());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setColor(new Color(COLOR::COLOR_BLUE)));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertEquals('FF0000FF', $object->getColor()->getARGB());
    }

    /**
     * Test get/set dash style.
     */
    public function testSetGetDashStyle(): void
    {
        $object = new Border();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle());
        $this->assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle(''));
        $this->assertEquals(Border::DASH_SOLID, $object->getDashStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setDashStyle(BORDER::DASH_DASH));
        $this->assertEquals(Border::DASH_DASH, $object->getDashStyle());
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $object = new Border();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set line style.
     */
    public function testSetGetLineStyle(): void
    {
        $object = new Border();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle());
        $this->assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle(''));
        $this->assertEquals(Border::LINE_SINGLE, $object->getLineStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineStyle(BORDER::LINE_DOUBLE));
        $this->assertEquals(Border::LINE_DOUBLE, $object->getLineStyle());
    }

    /**
     * Test get/set line width.
     */
    public function testSetGetLineWidth(): void
    {
        $object = new Border();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineWidth());
        $this->assertEquals(1, $object->getLineWidth());
        $value = mt_rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->setLineWidth($value));
        $this->assertEquals($value, $object->getLineWidth());
    }
}
