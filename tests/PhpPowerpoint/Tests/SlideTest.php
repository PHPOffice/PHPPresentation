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

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\Slide;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;

/**
 * Test class for PhpPowerpoint
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\PhpPowerpoint
 */
class SlideTest extends \PHPUnit_Framework_TestCase
{
    public function testParent()
    {
        $object = new Slide();
        $this->assertNull($object->getParent());
        
        $oPhpPowerpoint = new PhpPowerpoint();
        $object = new Slide($oPhpPowerpoint);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\PhpPowerpoint', $object->getParent());
    }
    
    public function testSlideLayout()
    {
        $object = new Slide();
        $this->assertEquals(Slide\Layout::BLANK, $object->getSlideLayout());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->setSlideLayout());
        $this->assertEquals(Slide\Layout::BLANK, $object->getSlideLayout());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->setSlideLayout(Slide\Layout::TITLE_SLIDE));
        $this->assertEquals(Slide\Layout::TITLE_SLIDE, $object->getSlideLayout());
    }
    
    public function testSlideMasterId()
    {
        $value = rand(1, 100);
        
        $object = new Slide();
        $this->assertEquals(1, $object->getSlideMasterId());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->setSlideMasterId());
        $this->assertEquals(1, $object->getSlideMasterId());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->setSlideMasterId($value));
        $this->assertEquals($value, $object->getSlideMasterId());
    }
}
