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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\Comment;

/**
 * Test class for Chart element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Comment
 */
class CommentTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Comment();

        $this->assertNull($object->getAuthor());
        $this->assertNull($object->getText());
        $this->assertInternalType('int', $object->getDate());
        $this->assertNull($object->getHeight());
        $this->assertNull($object->getWidth());
    }

    public function testGetSetAuthor()
    {
        $object = new Comment();

        $oStub = $this->getMockBuilder('PhpOffice\PhpPresentation\Shape\Comment\Author')->getMock();

        $this->assertNull($object->getAuthor());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment', $object->setAuthor($oStub));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment\\Author', $object->getAuthor());
    }

    public function testGetSetDate()
    {
        $expectedDate = time();

        $object = new Comment();
        $this->assertInternalType('int', $object->getDate());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment', $object->setDate($expectedDate));
        $this->assertEquals($expectedDate, $object->getDate());
        $this->assertInternalType('int', $object->getDate());
    }

    public function testGetSetText()
    {
        $expectedText = 'AABBCCDD';

        $object = new Comment();
        $this->assertNull($object->getText());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment', $object->setText($expectedText));
        $this->assertEquals($expectedText, $object->getText());
    }

    public function testGetSetHeightAndWidtg()
    {
        $object = new Comment();
        $this->assertNull($object->getHeight());
        $this->assertNull($object->getWidth());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment', $object->setHeight(1));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment', $object->setWidth(1));
        $this->assertNull($object->getHeight());
        $this->assertNull($object->getWidth());
    }
}
