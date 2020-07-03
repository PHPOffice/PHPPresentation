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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart\Type;

use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;

/**
 * Test class for Scatter element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter
 */
class AbstractTest extends \PHPUnit_Framework_TestCase
{
    public function testAxis()
    {
        $object = new Scatter();

        $this->assertTrue($object->hasAxisX());
        $this->assertTrue($object->hasAxisY());
    }

    public function testHashIndex()
    {
        $object = new Scatter();
        $value = rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testData()
    {
        $stub = $this->getMockForAbstractClass('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\AbstractType');
        $this->assertEmpty($stub->getData());
        $this->assertInternalType('array', $stub->getData());

        $arraySeries = array(
            new Series(),
            new Series()
        );
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\AbstractType', $stub->setData($arraySeries));
        $this->assertInternalType('array', $stub->getData());
        $this->assertCount(count($arraySeries), $stub->getData());
    }

    public function testClone()
    {
        $arraySeries = array(
            new Series(),
            new Series(),
            new Series(),
            new Series(),
        );

        $stub = $this->getMockForAbstractClass('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\AbstractType');
        $stub->setData($arraySeries);
        $clone = clone $stub;

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\AbstractType', $clone);
        $this->assertInternalType('array', $stub->getData());
        $this->assertCount(count($arraySeries), $stub->getData());
    }
}
