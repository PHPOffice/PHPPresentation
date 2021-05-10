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

use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class BulletTest extends TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new Bullet();
        static::assertEquals(Bullet::TYPE_NONE, $object->getBulletType());
        static::assertEquals('Calibri', $object->getBulletFont());
        static::assertEquals('-', $object->getBulletChar());
        static::assertEquals(Bullet::NUMERIC_DEFAULT, $object->getBulletNumericStyle());
        static::assertEquals(1, $object->getBulletNumericStartAt());
    }

    /**
     * Test get/set bullet char
     */
    public function testSetGetBulletChar()
    {
        $object = new Bullet();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletChar());
        static::assertEquals('-', $object->getBulletChar());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletChar('a'));
        static::assertEquals('a', $object->getBulletChar());
    }

    /**
     * Test get/set bullet color
     */
    public function testSetGetBulletColor()
    {
        $object = new Bullet();

        $expectedARGB = '01234567';

        // default
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getBulletColor());
        static::assertEquals(Color::COLOR_BLACK, $object->getBulletColor()->getARGB());


        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletColor(new Color($expectedARGB)));
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getBulletColor());
        static::assertEquals($expectedARGB, $object->getBulletColor()->getARGB());
    }

    /**
     * Test get/set bullet font
     */
    public function testSetGetBulletFont()
    {
        $object = new Bullet();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletFont());
        static::assertEquals('Calibri', $object->getBulletFont());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletFont(''));
        static::assertEquals('Calibri', $object->getBulletFont());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletFont('Arial'));
        static::assertEquals('Arial', $object->getBulletFont());
    }

    /**
     * Test get/set bullet numeric start at
     */
    public function testSetGetBulletNumericStartAt()
    {
        $object = new Bullet();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletNumericStartAt());
        static::assertEquals(1, $object->getBulletNumericStartAt());
        $value = mt_rand(1, 100);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletNumericStartAt($value));
        static::assertEquals($value, $object->getBulletNumericStartAt());
    }

    /**
     * Test get/set bullet numeric style
     */
    public function testSetGetBulletNumericStyle()
    {
        $object = new Bullet();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletNumericStyle());
        static::assertEquals(Bullet::NUMERIC_DEFAULT, $object->getBulletNumericStyle());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletNumericStyle(Bullet::NUMERIC_ALPHALCPARENBOTH));
        static::assertEquals(Bullet::NUMERIC_ALPHALCPARENBOTH, $object->getBulletNumericStyle());
    }

    /**
     * Test get/set bullet type
     */
    public function testSetGetBulletType()
    {
        $object = new Bullet();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletType());
        static::assertEquals(Bullet::TYPE_NONE, $object->getBulletType());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Bullet', $object->setBulletType(Bullet::TYPE_BULLET));
        static::assertEquals(Bullet::TYPE_BULLET, $object->getBulletType());
    }

    /**
     * Test get/set has index
     */
    public function testSetGetHashIndex()
    {
        $object = new Bullet();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        static::assertEquals($value, $object->getHashIndex());
    }
}
