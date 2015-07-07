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

use PhpOffice\PhpPowerpoint\Slide\Transition;

/**
 * Test class for PhpPowerpoint
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Slide\Transition
 */
class TransitionTest extends \PHPUnit_Framework_TestCase
{
    public function testSpeed()
    {
        $object = new Transition();
        $this->assertNull($object->getSpeed());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setSpeed());
        $this->assertEquals(Transition::SPEED_MEDIUM, $object->getSpeed());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setSpeed(Transition::SPEED_FAST));
        $this->assertEquals(Transition::SPEED_FAST, $object->getSpeed());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setSpeed(rand(1, 1000)));
        $this->assertNull($object->getSpeed());
    }

    public function testManualTrigger()
    {
        $object = new Transition();
        $this->assertFalse($object->hasManualTrigger());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setManualTrigger());
        $this->assertFalse($object->hasManualTrigger());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setManualTrigger(true));
        $this->assertTrue($object->hasManualTrigger());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setManualTrigger(null));
        $this->assertTrue($object->hasManualTrigger());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setManualTrigger(false));
        $this->assertFalse($object->hasManualTrigger());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setManualTrigger(null));
        $this->assertFalse($object->hasManualTrigger());
    }

    public function testTimeTrigger()
    {
        $object = new Transition();
        $this->assertFalse($object->hasTimeTrigger());
        $this->assertNull($object->getAdvanceTimeTrigger());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setTimeTrigger());
        $this->assertFalse($object->hasTimeTrigger());
        $this->assertNull($object->getAdvanceTimeTrigger());
        $value = rand(1, 1000);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setTimeTrigger(true, $value));
        $this->assertTrue($object->hasTimeTrigger());
        $this->assertEquals($value, $object->getAdvanceTimeTrigger());
        $value = rand(1, 1000);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setTimeTrigger(null, $value));
        $this->assertTrue($object->hasTimeTrigger());
        $this->assertEquals($value, $object->getAdvanceTimeTrigger());
        $value = rand(1, 1000);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setTimeTrigger(false, $value));
        $this->assertFalse($object->hasTimeTrigger());
        $this->assertNull($object->getAdvanceTimeTrigger());
        $value = rand(1, 1000);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setTimeTrigger(null, $value));
        $this->assertFalse($object->hasTimeTrigger());
        $this->assertNull($object->getAdvanceTimeTrigger());
    }

    public function testTransitionType()
    {
        $object = new Transition();
        $this->assertNull($object->getTransitionType());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setTransitionType());
        $this->assertNull($object->getTransitionType());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide\\Transition', $object->setTransitionType(Transition::TRANSITION_RANDOM));
        $this->assertEquals(Transition::TRANSITION_RANDOM, $object->getTransitionType());

    }
}
