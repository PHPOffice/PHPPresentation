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

namespace PhpOffice\PhpPowerpoint\Tests\Slide;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Slide;
use PhpOffice\PhpPowerpoint\Slide\Iterator;

/**
 * Test class for IOFactory
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\IOFactory
 */
class IteratorTest extends \PHPUnit_Framework_TestCase
{
    /**
     */
    public function testMethod()
    {
        $oPHPPowerPoint = new PhpPowerpoint();
        $oPHPPowerPoint->addSlide(new Slide());

        $object = new Iterator($oPHPPowerPoint);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->current());
        $this->assertEquals(0, $object->key());
        $this->assertNull($object->next());
        $this->assertEquals(1, $object->key());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->current());
        $this->assertTrue($object->valid());
        $this->assertNull($object->next());
        $this->assertFalse($object->valid());
        $this->assertNull($object->rewind());
        $this->assertEquals(0, $object->key());
    }
}
