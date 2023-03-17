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
use PhpOffice\PhpPresentation\Slide\Animation;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Animation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Slide\Animation
 */
class AnimationTest extends TestCase
{
    public function testShape(): void
    {
        /** @var AbstractShape $oStub */
        $oStub = $this->getMockForAbstractClass(AbstractShape::class);

        $object = new Animation();

        $this->assertIsArray($object->getShapeCollection());
        $this->assertCount(0, $object->getShapeCollection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Animation', $object->addShape($oStub));
        $this->assertIsArray($object->getShapeCollection());
        $this->assertCount(1, $object->getShapeCollection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Animation', $object->setShapeCollection());
        $this->assertIsArray($object->getShapeCollection());
        $this->assertCount(0, $object->getShapeCollection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Animation', $object->setShapeCollection([$oStub]));
        $this->assertIsArray($object->getShapeCollection());
        $this->assertCount(1, $object->getShapeCollection());
    }
}
