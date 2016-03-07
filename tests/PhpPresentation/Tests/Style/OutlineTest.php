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

use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Style\Outline
 */
class OutlineTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Outline();
        $this->assertNull($object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
    }

    /**
     * Test get/set fill
     */
    public function testSetGetFill()
    {
        $object = new Outline();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Outline', $object->setFill(new Fill()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
    }

    /**
     * Test get/set width
     */
    public function testSetGetWidth()
    {
        $object = new Outline();
        $this->assertNull($object->getWidth());
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Outline', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Outline', $object->setWidth(1.5));
        $this->assertEquals(1, $object->getWidth());
    }
}
