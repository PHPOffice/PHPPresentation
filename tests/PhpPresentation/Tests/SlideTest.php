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

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\AbstractBackground;
use PhpOffice\PhpPresentation\Slide\Animation;
use PhpOffice\PhpPresentation\Slide\Transition;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class SlideTest extends TestCase
{
    public function testExtents(): void
    {
        $object = new Slide();
        $this->assertNotNull($object->getExtentX());

        $object = new Slide();
        $this->assertNotNull($object->getExtentY());
    }

    public function testOffset(): void
    {
        $object = new Slide();
        $this->assertNotNull($object->getOffsetX());

        $object = new Slide();
        $this->assertNotNull($object->getOffsetY());
    }

    public function testParent(): void
    {
        $object = new Slide();
        $this->assertNull($object->getParent());

        $oPhpPresentation = new PhpPresentation();
        $object = new Slide($oPhpPresentation);
        $this->assertInstanceOf(PhpPresentation::class, $object->getParent());
    }

    public function testSlideMasterId(): void
    {
        $value = mt_rand(1, 100);
        $object = new Slide();
        $this->assertEquals(1, $object->getSlideMasterId());
        $this->assertInstanceOf(Slide::class, $object->setSlideMasterId());
        $this->assertEquals(1, $object->getSlideMasterId());
        $this->assertInstanceOf(Slide::class, $object->setSlideMasterId($value));
        $this->assertEquals($value, $object->getSlideMasterId());
    }

    public function testAnimations(): void
    {
        /** @var Animation $oStub */
        $oStub = $this->getMockForAbstractClass(Animation::class);

        $object = new Slide();
        $this->assertIsArray($object->getAnimations());
        $this->assertCount(0, $object->getAnimations());
        $this->assertInstanceOf(Slide::class, $object->addAnimation($oStub));
        $this->assertIsArray($object->getAnimations());
        $this->assertCount(1, $object->getAnimations());
        $this->assertInstanceOf(Slide::class, $object->setAnimations());
        $this->assertIsArray($object->getAnimations());
        $this->assertCount(0, $object->getAnimations());
        $this->assertInstanceOf(Slide::class, $object->setAnimations([$oStub]));
        $this->assertIsArray($object->getAnimations());
        $this->assertCount(1, $object->getAnimations());
    }

    public function testBackground(): void
    {
        /** @var AbstractBackground $oStub */
        $oStub = $this->getMockForAbstractClass(AbstractBackground::class);

        $object = new Slide();
        $this->assertNull($object->getBackground());
        $this->assertInstanceOf(Slide::class, $object->setBackground($oStub));
        $this->assertInstanceOf(AbstractBackground::class, $object->getBackground());
        $this->assertInstanceOf(Slide::class, $object->setBackground());
        $this->assertNull($object->getBackground());
    }

    public function testGroup(): void
    {
        $object = new Slide();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Group', $object->createGroup());
    }

    public function testName(): void
    {
        $object = new Slide();
        $this->assertNull($object->getName());
        $this->assertInstanceOf(Slide::class, $object->setName('AAAA'));
        $this->assertEquals('AAAA', $object->getName());
        $this->assertInstanceOf(Slide::class, $object->setName());
        $this->assertNull($object->getName());
    }

    public function testTransition(): void
    {
        $object = new Slide();
        $oTransition = new Transition();
        $this->assertNull($object->getTransition());
        $this->assertInstanceOf(Slide::class, $object->setTransition());
        $this->assertNull($object->getTransition());
        $this->assertInstanceOf(Slide::class, $object->setTransition($oTransition));
        $this->assertInstanceOf(Transition::class, $object->getTransition());
        $this->assertInstanceOf(Slide::class, $object->setTransition(null));
        $this->assertNull($object->getTransition());
    }

    public function testVisible(): void
    {
        $object = new Slide();
        $this->assertTrue($object->isVisible());
        $this->assertInstanceOf(Slide::class, $object->setIsVisible(false));
        $this->assertFalse($object->isVisible());
        $this->assertInstanceOf(Slide::class, $object->setIsVisible());
        $this->assertTrue($object->isVisible());
    }
}
