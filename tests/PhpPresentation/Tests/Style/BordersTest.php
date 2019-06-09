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
use PhpOffice\PhpPresentation\Style\Borders;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class BordersTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Borders();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getBottom());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getLeft());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getRight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getTop());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getDiagonalDown());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getDiagonalUp());
        $this->assertEquals(Border::LINE_NONE, $object->getDiagonalDown()->getLineStyle());
        $this->assertEquals(Border::LINE_NONE, $object->getDiagonalUp()->getLineStyle());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Borders();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set hash code
     */
    public function testGetHashCode()
    {
        $object = new Borders();
        $this->assertEquals(
            md5(
                $object->getLeft()->getHashCode() .
                $object->getRight()->getHashCode() .
                $object->getTop()->getHashCode() .
                $object->getBottom()->getHashCode() .
                $object->getDiagonalUp()->getHashCode() .
                $object->getDiagonalDown()->getHashCode() .
                get_class($object)
            ),
            $object->getHashCode()
        );
    }
}
