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
 * @copyright   2010-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\Slide;
use PhpOffice\PhpPowerpoint\Shape\Hyperlink;
use PhpOffice\PhpPowerpoint\Shape\RichText;
use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Style\Shadow;

/**
 * Test class for Autoloader
 */
class AbstractShapeTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Register
     */
    public function testConstruct()
    {
        $object = new RichText();

        $this->assertNull($object->getSlide());
        $this->assertEquals(0, $object->getOffsetX());
        $this->assertEquals(0, $object->getOffsetY());
        $this->assertEquals(0, $object->getHeight());
        $this->assertEquals(0, $object->getRotation());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Border', $object->getBorder());
        $this->assertEquals(Border::LINE_NONE, $object->getBorder()->getLineStyle());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Fill', $object->getFill());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->getShadow());
    }

    public function testFill()
    {
        $object = new RichText();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setFill());
        $this->assertNull($object->getFill());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setFill(new Fill()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Fill', $object->getFill());
    }

    public function testHeight()
    {
        $object = new RichText();

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setHeight());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setHeight($value));
        $this->assertEquals($value, $object->getHeight());
    }

    public function testHyperlink()
    {
        $object = new RichText();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setHyperlink());
        $this->assertFalse($object->hasHyperlink());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Hyperlink', $object->getHyperlink());
        $this->assertTrue($object->hasHyperlink());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setHyperlink(new Hyperlink('http://www.google.fr')));
        $this->assertTrue($object->hasHyperlink());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Hyperlink', $object->getHyperlink());
        $this->assertTrue($object->hasHyperlink());
    }

    public function testOffsetX()
    {
        $object = new RichText();

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setOffsetX());
        $this->assertEquals(0, $object->getOffsetX());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setOffsetX($value));
        $this->assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY()
    {
        $object = new RichText();

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setOffsetY());
        $this->assertEquals(0, $object->getOffsetY());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setOffsetY($value));
        $this->assertEquals($value, $object->getOffsetY());
    }

    public function testRotation()
    {
        $object = new RichText();

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setRotation());
        $this->assertEquals(0, $object->getRotation());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setRotation($value));
        $this->assertEquals($value, $object->getRotation());
    }

    public function testShadow()
    {
        $object = new RichText();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setShadow());
        $this->assertNull($object->getShadow());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setShadow(new Shadow()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Style\\Shadow', $object->getShadow());
    }

    public function testSlide()
    {
        $object = new RichText();
        $oSlide1 = new Slide();
        $oSlide2 = new Slide();
        $oSlide3 = new Slide();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setSlide());
        $this->assertNull($object->getSlide());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setSlide($oSlide1, true));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->getSlide());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setSlide($oSlide2, true));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->getSlide());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setSlide($oSlide3, true));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->getSlide());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setSlide($oSlide3, true));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Slide', $object->getSlide());
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage A \PhpOffice\PhpPowerpoint\Slide has already been assigned. Shapes can only exist on one \PhpOffice\PhpPowerpoint\Slide.
     */
    public function testSlideException()
    {
        $object = new RichText();
        $object->setSlide(new Slide());
        $object->setSlide(new Slide());
    }

    public function testWidth()
    {
        $object = new RichText();

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setWidth());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
    }

    public function testWidthAndHeight()
    {
        $object = new RichText();

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setWidthAndHeight());
        $this->assertEquals(0, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setWidthAndHeight($value));
        $this->assertEquals($value, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\AbstractShape', $object->setWidthAndHeight($value, $value));
        $this->assertEquals($value, $object->getWidth());
        $this->assertEquals($value, $object->getHeight());
    }
}
