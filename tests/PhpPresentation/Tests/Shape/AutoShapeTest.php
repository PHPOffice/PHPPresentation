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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Style\Outline;
use PHPUnit\Framework\TestCase;

class AutoShapeTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new AutoShape();

        $this->assertEquals(AutoShape::TYPE_HEART, $object->getType());
        $this->assertEquals('', $object->getText());
        $this->assertInstanceOf(Outline::class, $object->getOutline());
        $this->assertIsString($object->getHashCode());
    }

    public function testOutline(): void
    {
        /** @var Outline $mock */
        $mock = $this->getMockBuilder(Outline::class)->getMock();

        $object = new AutoShape();
        $this->assertInstanceOf(Outline::class, $object->getOutline());
        $this->assertInstanceOf(AutoShape::class, $object->setOutline($mock));
        $this->assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testText(): void
    {
        $object = new AutoShape();

        $this->assertEquals('', $object->getText());
        $this->assertInstanceOf(AutoShape::class, $object->setText('Text'));
        $this->assertEquals('Text', $object->getText());
    }

    public function testType(): void
    {
        $object = new AutoShape();

        $this->assertEquals(AutoShape::TYPE_HEART, $object->getType());
        $this->assertInstanceOf(AutoShape::class, $object->setType(AutoShape::TYPE_HEXAGON));
        $this->assertEquals(AutoShape::TYPE_HEXAGON, $object->getType());
    }
}
