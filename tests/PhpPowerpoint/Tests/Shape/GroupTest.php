<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Shape;

use PhpOffice\PhpPowerpoint\Shape\Group;
use PhpOffice\PhpPowerpoint\Shape\Line;

/**
 * Test class for Group element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Group
 */
class GroupTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Group();
        
        $this->assertEquals(0, $object->getOffsetX());
        $this->assertEquals(0, $object->getOffsetY());
        $this->assertEquals(0, $object->getExtentX());
        $this->assertEquals(0, $object->getExtentY());
        $this->assertEquals(0, $object->getShapeCollection()->count());
    }
    
    public function testExtentsAndOffsetsForOneShape()
    {
        // We record initial values here because
        // PhpOffice\PhpPowerpoint\Shape\Line subtracts the offsets
        // from the extents to produce a raw width and height.
        $offsetX = 100;
        $offsetY = 100;
        $extentX = 1000;
        $extentY = 450;

        $line1  = new Line($offsetX, $offsetY, $extentX, $extentY);
        $object = new Group();
        
        $object->addShape($line1);
        
        $this->assertEquals($offsetX, $object->getOffsetX());
        $this->assertEquals($offsetY, $object->getOffsetY());
        $this->assertEquals($extentX, $object->getExtentX());
        $this->assertEquals($extentY, $object->getExtentY());
    }

    public function testExtentsAndOffsetsForTwoShapes()
    {
        // Since Groups and Slides cache offsets and extents on first
        // calculation, this test is separate from the above.
        // Should the calculation be performed every GET, this test can be
        // combined with the above.
        $offsetX = 100;
        $offsetY = 100;
        $extentX = 1000;
        $extentY = 450;
        $increase = 50;

        $line1  = new Line($offsetX, $offsetY, $extentX, $extentY);
        $line2 = new Line(
            $offsetX+$increase,
            $offsetY+$increase,
            $extentX+$increase,
            $extentY+$increase
        );

        $object = new Group();
        
        $object->addShape($line1);
        $object->addShape($line2);
        
        $this->assertEquals($offsetX, $object->getOffsetX());
        $this->assertEquals($offsetY, $object->getOffsetY());
        $this->assertEquals($extentX+$increase, $object->getExtentX());
        $this->assertEquals($extentY+$increase, $object->getExtentY());
    }
}
