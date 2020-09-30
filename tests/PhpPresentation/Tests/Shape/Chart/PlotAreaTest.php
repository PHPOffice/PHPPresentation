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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\PlotArea;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PlotArea element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart\PlotArea
 */
class PlotAreaTest extends TestCase
{
    public function testConstruct()
    {
        $object = new PlotArea();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->getAxisX());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->getAxisY());
    }

    public function testHashIndex()
    {
        $object = new PlotArea();
        $value = mt_rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testHeight()
    {
        $object = new PlotArea();
        $value = mt_rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setHeight());
        $this->assertEquals(0, $object->getHeight());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setHeight($value));
        $this->assertEquals($value, $object->getHeight());
    }

    public function testOffsetX()
    {
        $object = new PlotArea();
        $value = mt_rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setOffsetX());
        $this->assertEquals(0, $object->getOffsetX());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setOffsetX($value));
        $this->assertEquals($value, $object->getOffsetX());
    }

    public function testOffsetY()
    {
        $object = new PlotArea();
        $value = mt_rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setOffsetY());
        $this->assertEquals(0, $object->getOffsetY());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setOffsetY($value));
        $this->assertEquals($value, $object->getOffsetY());
    }

    public function testType()
    {
        $object = new PlotArea();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setType(new Bar3D()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\AbstractType', $object->getType());
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
        $value = mt_rand(0, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setWidth());
        $this->assertEquals(0, $object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\PlotArea', $object->setWidth($value));
        $this->assertEquals($value, $object->getWidth());
    }
}
