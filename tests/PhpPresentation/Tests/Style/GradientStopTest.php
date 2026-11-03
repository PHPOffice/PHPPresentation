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

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\GradientStop;
use PHPUnit\Framework\TestCase;

class GradientStopTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new GradientStop();
        self::assertEquals(GradientStop::POSITION_START, $object->getPosition());
        self::assertEquals(Color::COLOR_BLACK, $object->getColor()->getARGB());
    }

    public function testConstructWithArguments(): void
    {
        $color = new Color(Color::COLOR_RED);
        $object = new GradientStop(50000, $color);
        self::assertEquals(50000, $object->getPosition());
        self::assertSame($color, $object->getColor());
    }

    public function testSetGetPosition(): void
    {
        $object = new GradientStop();
        self::assertInstanceOf(GradientStop::class, $object->setPosition(25000));
        self::assertEquals(25000, $object->getPosition());
    }

    public function testSetGetColor(): void
    {
        $object = new GradientStop();
        $color = new Color(Color::COLOR_BLUE);
        self::assertInstanceOf(GradientStop::class, $object->setColor($color));
        self::assertSame($color, $object->getColor());
    }

    public function testHashCode(): void
    {
        $object = new GradientStop(50000, new Color(Color::COLOR_RED));
        $same = new GradientStop(50000, new Color(Color::COLOR_RED));
        $other = new GradientStop(50000, new Color(Color::COLOR_BLUE));

        self::assertEquals($object->getHashCode(), $same->getHashCode());
        self::assertNotEquals($object->getHashCode(), $other->getHashCode());
    }
}
