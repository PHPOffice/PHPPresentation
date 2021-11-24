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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\Exception\ShapeContainerAlreadyAssignedException;
use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Autoloader.
 */
class AbstractShapeTest extends TestCase
{
    /**
     * Register.
     */
    public function testConstruct(): void
    {
        $object = new RichText();

        $this->assertEquals(0, $object->getOffsetX());
        $this->assertEquals(0, $object->getOffsetY());
        $this->assertEquals(0, $object->getHeight());
        $this->assertEquals(0, $object->getRotation());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf(Border::class, $object->getBorder());
        $this->assertEquals(Border::LINE_NONE, $object->getBorder()->getLineStyle());
        $this->assertInstanceOf(Fill::class, $object->getFill());
        $this->assertInstanceOf(Shadow::class, $object->getShadow());
    }

    public function testFill(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(AbstractShape::class, $object->setFill());
        $this->assertNull($object->getFill());
        $this->assertInstanceOf(AbstractShape::class, $object->setFill(new Fill()));
        $this->assertInstanceOf(Fill::class, $object->getFill());
    }

    public function testHeight(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        $this->assertInstanceOf(AbstractShape::class, $object->setHeight());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf(AbstractShape::class, $object->setHeight($value));
        $this->assertEquals($value, $object->getHeight());
    }

    public function testHyperlink(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(AbstractShape::class, $object->setHyperlink());
        $this->assertFalse($object->hasHyperlink());
        $this->assertInstanceOf(Hyperlink::class, $object->getHyperlink());
        $this->assertTrue($object->hasHyperlink());
        $this->assertInstanceOf(AbstractShape::class, $object->setHyperlink(new Hyperlink('http://www.google.fr')));
        $this->assertTrue($object->hasHyperlink());
        $this->assertInstanceOf(Hyperlink::class, $object->getHyperlink());
        $this->assertTrue($object->hasHyperlink());
    }

    public function testOffsetX(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        $this->assertInstanceOf(AbstractShape::class, $object->setOffsetX());
        $this->assertEquals(0, $object->getOffsetX());
        $this->assertInstanceOf(AbstractShape::class, $object->setOffsetX($value));
        $this->assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        $this->assertInstanceOf(AbstractShape::class, $object->setOffsetY());
        $this->assertEquals(0, $object->getOffsetY());
        $this->assertInstanceOf(AbstractShape::class, $object->setOffsetY($value));
        $this->assertEquals($value, $object->getOffsetY());
    }

    public function testRotation(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        $this->assertInstanceOf(AbstractShape::class, $object->setRotation());
        $this->assertEquals(0, $object->getRotation());
        $this->assertInstanceOf(AbstractShape::class, $object->setRotation($value));
        $this->assertEquals($value, $object->getRotation());
    }

    public function testShadow(): void
    {
        $object = new RichText();

        $this->assertInstanceOf(AbstractShape::class, $object->setShadow());
        $this->assertNull($object->getShadow());
        $this->assertInstanceOf(AbstractShape::class, $object->setShadow(new Shadow()));
        $this->assertInstanceOf(Shadow::class, $object->getShadow());
    }

    public function testWidth(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        $this->assertInstanceOf(AbstractShape::class, $object->setWidth());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf(AbstractShape::class, $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
    }

    public function testWidthAndHeight(): void
    {
        $object = new RichText();

        $value = mt_rand(1, 100);
        $this->assertInstanceOf(AbstractShape::class, $object->setWidthAndHeight());
        $this->assertEquals(0, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf(AbstractShape::class, $object->setWidthAndHeight($value));
        $this->assertEquals($value, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf(AbstractShape::class, $object->setWidthAndHeight($value, $value));
        $this->assertEquals($value, $object->getWidth());
        $this->assertEquals($value, $object->getHeight());
    }

    public function testPlaceholder(): void
    {
        $object = new RichText();
        $this->assertFalse($object->isPlaceholder(), 'Standard Shape should not be a placeholder object');
        $this->assertNull($object->getPlaceholder());
        $this->assertInstanceOf(
            AbstractShape::class,
            $object->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_TITLE))
        );
        $this->assertTrue($object->isPlaceholder());
        $this->assertInstanceOf(Placeholder::class, $object->getPlaceholder());
        $this->assertEquals('title', $object->getPlaceholder()->getType());

        $object = new RichText();
        $this->assertFalse($object->isPlaceholder(), 'Standard Shape should not be a placeholder object');
        $placeholder = new Placeholder(Placeholder::PH_TYPE_TITLE);
        $placeholder->setType(Placeholder::PH_TYPE_SUBTITLE);
        $this->assertInstanceOf(AbstractShape::class, $object->setPlaceHolder($placeholder));
        $this->assertTrue($object->isPlaceholder());
        $this->assertInstanceOf(Placeholder::class, $object->getPlaceholder());
        $this->assertEquals('subTitle', $object->getPlaceholder()->getType());
    }

    public function testContainer(): void
    {
        $object = new RichText();
        $object2 = new RichText();
        $object2->setWrap(RichText::WRAP_NONE);
        $oSlide = new Slide();
        $oSlide->addShape($object2);

        $this->assertNull($object->getContainer());
        $this->assertInstanceOf(AbstractShape::class, $object->setContainer($oSlide));
        $this->assertInstanceOf(Slide::class, $object->getContainer());
        $this->assertInstanceOf(AbstractShape::class, $object->setContainer(null, true));
        $this->assertNull($object->getContainer());
    }

    public function testContainerException(): void
    {
        $object = new RichText();
        $oSlide = new Slide();

        $this->assertNull($object->getContainer());
        $this->assertInstanceOf(AbstractShape::class, $object->setContainer($oSlide));
        $this->assertInstanceOf(Slide::class, $object->getContainer());
        $this->expectException(ShapeContainerAlreadyAssignedException::class);
        $this->expectExceptionMessage('The shape PhpOffice\PhpPresentation\AbstractShape has already a container assigned');
        $object->setContainer(null);
    }
}
