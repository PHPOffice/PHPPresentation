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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\Slide\Transition;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Slide\Transition
 */
class TransitionTest extends TestCase
{
    public function testSpeed(): void
    {
        $object = new Transition();
        self::assertNull($object->getSpeed());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setSpeed());
        self::assertEquals(Transition::SPEED_MEDIUM, $object->getSpeed());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setSpeed(Transition::SPEED_FAST));
        self::assertEquals(Transition::SPEED_FAST, $object->getSpeed());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setSpeed('notagoodvalue'));
        self::assertNull($object->getSpeed());
    }

    public function testManualTrigger(): void
    {
        $object = new Transition();
        self::assertFalse($object->hasManualTrigger());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setManualTrigger());
        self::assertFalse($object->hasManualTrigger());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setManualTrigger(true));
        self::assertTrue($object->hasManualTrigger());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setManualTrigger(false));
        self::assertFalse($object->hasManualTrigger());
    }

    public function testTimeTrigger(): void
    {
        $object = new Transition();
        self::assertFalse($object->hasTimeTrigger());
        self::assertNull($object->getAdvanceTimeTrigger());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setTimeTrigger());
        self::assertFalse($object->hasTimeTrigger());
        self::assertNull($object->getAdvanceTimeTrigger());
        $value = mt_rand(1, 1000);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setTimeTrigger(true, $value));
        self::assertTrue($object->hasTimeTrigger());
        self::assertEquals($value, $object->getAdvanceTimeTrigger());
        $value = mt_rand(1, 1000);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setTimeTrigger(false, $value));
        self::assertFalse($object->hasTimeTrigger());
        self::assertNull($object->getAdvanceTimeTrigger());
    }

    public function testTransitionType(): void
    {
        $object = new Transition();
        self::assertNull($object->getTransitionType());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setTransitionType());
        self::assertNull($object->getTransitionType());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Transition', $object->setTransitionType(Transition::TRANSITION_RANDOM));
        self::assertEquals(Transition::TRANSITION_RANDOM, $object->getTransitionType());
    }
}
