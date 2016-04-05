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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\Slide\Animation;

/**
 * Test class for Animation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Slide\Animation
 */
class AnimationTest extends \PHPUnit_Framework_TestCase
{
    public function testShape()
    {
        $oStub = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\AbstractShape');

        $object = new Animation();

        $this->assertInternalType('array', $object->getShapeCollection());
        $this->assertCount(0, $object->getShapeCollection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Animation', $object->addShape($oStub));
        $this->assertInternalType('array', $object->getShapeCollection());
        $this->assertCount(1, $object->getShapeCollection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Animation', $object->setShapeCollection());
        $this->assertInternalType('array', $object->getShapeCollection());
        $this->assertCount(0, $object->getShapeCollection());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Animation', $object->setShapeCollection(array($oStub)));
        $this->assertInternalType('array', $object->getShapeCollection());
        $this->assertCount(1, $object->getShapeCollection());
    }
}
