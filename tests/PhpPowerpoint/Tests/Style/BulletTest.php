<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\Style\Bullet;
use PhpOffice\PhpPowerpoint\Style\Color;

/**
 * Test class for PhpPowerpoint
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\PhpPowerpoint
 */
class BulletTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct ()
    {
        $object = new Bullet();
        $this->assertEquals(Bullet::TYPE_NONE, $object->getBulletType());
        $this->assertEquals('Calibri', $object->getBulletFont());
        $this->assertEquals('-', $object->getBulletChar());
        $this->assertEquals(Bullet::NUMERIC_DEFAULT, $object->getBulletNumericStyle());
        $this->assertEquals(1, $object->getBulletNumericStartAt());
    }

    public function testSetGetBulletChar ()
    {
        $object = new Bullet();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletChar());
        $this->assertEquals('-', $object->getBulletChar());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletChar('a'));
        $this->assertEquals('a', $object->getBulletChar());
    }
    public function testSetGetBulletFont ()
    {
        $object = new Bullet();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletFont());
        $this->assertEquals('Calibri', $object->getBulletFont());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletFont(''));
        $this->assertEquals('Calibri', $object->getBulletFont());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletFont('Arial'));
        $this->assertEquals('Arial', $object->getBulletFont());
    }
    public function testSetGetBulletNumericStartAt ()
    {
        $object = new Bullet();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletNumericStartAt());
        $this->assertEquals(1, $object->getBulletNumericStartAt());
        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletNumericStartAt($value));
        $this->assertEquals($value, $object->getBulletNumericStartAt());
    }
    public function testSetGetBulletNumericStyle ()
    {
        $object = new Bullet();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletNumericStyle());
        $this->assertEquals(Bullet::NUMERIC_DEFAULT, $object->getBulletNumericStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletNumericStyle(Bullet::NUMERIC_ALPHALCPARENBOTH));
        $this->assertEquals(Bullet::NUMERIC_ALPHALCPARENBOTH, $object->getBulletNumericStyle());
    }
    
    public function testSetGetBulletType ()
    {
        $object = new Bullet();
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletType());
        $this->assertEquals(Bullet::TYPE_NONE, $object->getBulletType());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Bullet', $object->setBulletType(Bullet::TYPE_BULLET));
        $this->assertEquals(Bullet::TYPE_BULLET, $object->getBulletType());
    }
    
    public function testSetGetHashIndex ()
    {
        $object = new Bullet();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }
}
