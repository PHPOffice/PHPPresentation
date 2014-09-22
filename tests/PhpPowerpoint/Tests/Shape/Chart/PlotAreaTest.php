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
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Shape\Chart;

use PhpOffice\PhpPowerpoint\Shape\Chart;
use PhpOffice\PhpPowerpoint\Shape\Chart\Axis;
use PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D;

/**
 * Test class for PlotArea element
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea
 */
class PlotAreaTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new PlotArea();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Axis', $object->getAxisX());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Axis', $object->getAxisY());
    }

    public function testHashIndex()
    {
        $object = new PlotArea();
        $value = rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testHeight()
    {
        $object = new PlotArea();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setHeight());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setHeight($value));
        $this->assertEquals($value, $object->getHeight());
    }

    public function testOffsetX()
    {
        $object = new PlotArea();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setOffsetX());
        $this->assertEquals(0, $object->getOffsetX());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setOffsetX($value));
        $this->assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY()
    {
        $object = new PlotArea();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setOffsetY());
        $this->assertEquals(0, $object->getOffsetY());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setOffsetY($value));
        $this->assertEquals($value, $object->getOffsetY());
    }

    public function testType()
    {
        $object = new PlotArea();

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setType(new Bar3D()));
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\Type\\AbstractType', $object->getType());
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Chart type has not been set.
     */
    public function testTypeException()
    {
        $object = new PlotArea();
        $object->getType();
    }

    public function testWidth()
    {
        $object = new PlotArea();
        $value = rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setWidth());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\Shape\\Chart\\PlotArea', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
    }
}
