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

namespace PhpOffice\PhpPresentation\Tests\Shape\Table;

use PhpOffice\PhpPresentation\Shape\Table\Row;
use PhpOffice\PhpPresentation\Style\Fill;

/**
 * Test class for Row element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Row
 */
class RowTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testConstruct()
    {
        $object = new Row();
        $this->assertCount(1, $object->getCells());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());

        $value = rand(1, 100);
        $object = new Row($value);
        $this->assertCount($value, $object->getCells());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
    }

    public function testGetCell()
    {
        $object = new Row();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->getCell(0));
        $this->assertNull($object->getCell(1000, true));
    }

    /**
     * @expectedException \Exception
     * expectedExceptionMessage Cell number out of bounds.
     */
    public function testGetCellException()
    {
        $object = new Row();
        $object->getCell(1);
    }

    public function testNextCell()
    {
        $object = new Row(2);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Cell', $object->nextCell());
    }

    /**
     * @expectedException \Exception
     * expectedExceptionMessage Cell count out of bounds.
     */
    public function testNextCellException()
    {
        $object = new Row();
        $object->nextCell();
        $object->nextCell();
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Row();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testGetSetFill()
    {
        $object = new Row();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Row', $object->setFill(new Fill()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
    }

    public function testGetSetHeight()
    {
        $object = new Row();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Row', $object->setHeight());
        $this->assertEquals(0, $object->getHeight());

        $value = rand(1, 100);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Table\\Row', $object->setHeight($value));
        $this->assertEquals($value, $object->getHeight());
    }
}
