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

use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class BulletTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $object = new Bullet();
        self::assertEquals(Bullet::TYPE_NONE, $object->getBulletType());
        self::assertEquals('Calibri', $object->getBulletFont());
        self::assertEquals('-', $object->getBulletChar());
        self::assertEquals(Bullet::NUMERIC_DEFAULT, $object->getBulletNumericStyle());
        self::assertEquals(1, $object->getBulletNumericStartAt());
    }

    /**
     * Test get/set bullet char.
     */
    public function testSetGetBulletChar(): void
    {
        $object = new Bullet();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletChar());
        self::assertEquals('-', $object->getBulletChar());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletChar('a'));
        self::assertEquals('a', $object->getBulletChar());
    }

    /**
     * Test get/set bullet color.
     */
    public function testSetGetBulletColor(): void
    {
        $object = new Bullet();

        $expectedARGB = '01234567';

        // default
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getBulletColor());
        self::assertEquals(Color::COLOR_BLACK, $object->getBulletColor()->getARGB());

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletColor(new Color($expectedARGB)));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getBulletColor());
        self::assertEquals($expectedARGB, $object->getBulletColor()->getARGB());
    }

    /**
     * Test get/set bullet font.
     */
    public function testSetGetBulletFont(): void
    {
        $object = new Bullet();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletFont());
        self::assertEquals('Calibri', $object->getBulletFont());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletFont(''));
        self::assertEquals('Calibri', $object->getBulletFont());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletFont('Arial'));
        self::assertEquals('Arial', $object->getBulletFont());
    }

    /**
     * Test get/set bullet numeric start at.
     */
    public function testSetGetBulletNumericStartAt(): void
    {
        $object = new Bullet();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletNumericStartAt());
        self::assertEquals(1, $object->getBulletNumericStartAt());
        $value = mt_rand(1, 100);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletNumericStartAt($value));
        self::assertEquals($value, $object->getBulletNumericStartAt());
    }

    /**
     * Test get/set bullet numeric style.
     */
    public function testSetGetBulletNumericStyle(): void
    {
        $object = new Bullet();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletNumericStyle());
        self::assertEquals(Bullet::NUMERIC_DEFAULT, $object->getBulletNumericStyle());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletNumericStyle(Bullet::NUMERIC_ALPHALCPARENBOTH));
        self::assertEquals(Bullet::NUMERIC_ALPHALCPARENBOTH, $object->getBulletNumericStyle());
    }

    /**
     * Test get/set bullet type.
     */
    public function testSetGetBulletType(): void
    {
        $object = new Bullet();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletType());
        self::assertEquals(Bullet::TYPE_NONE, $object->getBulletType());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletType(Bullet::TYPE_BULLET));
        self::assertEquals(Bullet::TYPE_BULLET, $object->getBulletType());
    }

    /**
     * Test get/set has index.
     */
    public function testSetGetHashIndex(): void
    {
        $object = new Bullet();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }
}
