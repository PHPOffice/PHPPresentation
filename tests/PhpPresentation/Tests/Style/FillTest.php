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

use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class FillTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Fill();
        static::assertEquals(Fill::FILL_NONE, $object->getFillType());
        static::assertEquals(0, $object->getRotation());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getStartColor());
        static::assertEquals(Color::COLOR_WHITE, $object->getStartColor()->getARGB());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getEndColor());
        static::assertEquals(Color::COLOR_BLACK, $object->getEndColor()->getARGB());
    }

    /**
     * Test get/set end color
     */
    public function testSetGetEndColor()
    {
        $object = new Fill();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->setEndColor());
        static::assertNull($object->getEndColor());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->setEndColor(new Color(Color::COLOR_BLUE)));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getEndColor());
        static::assertEquals(Color::COLOR_BLUE, $object->getEndColor()->getARGB());
    }

    /**
     * Test get/set fill type
     */
    public function testSetGetFillType()
    {
        $object = new Fill();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->setFillType());
        static::assertEquals(Fill::FILL_NONE, $object->getFillType());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->setFillType(Fill::FILL_GRADIENT_LINEAR));
        static::assertEquals(Fill::FILL_GRADIENT_LINEAR, $object->getFillType());
    }

    /**
     * Test get/set rotation
     */
    public function testSetGetRotation()
    {
        $object = new Fill();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->setRotation());
        static::assertEquals(0, $object->getRotation());
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->setRotation($value));
        static::assertEquals($value, $object->getRotation());
    }

    /**
     * Test get/set start color
     */
    public function testSetGetStartColor()
    {
        $object = new Fill();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->setStartColor());
        static::assertNull($object->getStartColor());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->setStartColor(new Color(Color::COLOR_BLUE)));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getStartColor());
        static::assertEquals(Color::COLOR_BLUE, $object->getStartColor()->getARGB());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Fill();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        static::assertEquals($value, $object->getHashIndex());
    }
}
