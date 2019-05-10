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

namespace PhpOffice\PhpPresentation\Tests\Shape\Comment;

use PhpOffice\PhpPresentation\Shape\Comment\Author;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Author element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Comment\Author
 */
class AuthorTest extends TestCase
{
    public function testConstruct()
    {
        $object = new Author();

        $this->assertNull($object->getName());
        $this->assertNull($object->getIndex());
        $this->assertNull($object->getInitials());
    }

    public function testGetSetIndex()
    {
        $expectedVal = mt_rand(1, 100);

        $object = new Author();
        $this->assertNull($object->getIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment\\Author', $object->setIndex($expectedVal));
        $this->assertEquals($expectedVal, $object->getIndex());
    }

    public function testGetSetInitials()
    {
        $expectedVal = 'AABBCCDD';

        $object = new Author();
        $this->assertNull($object->getInitials());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment\\Author', $object->setInitials($expectedVal));
        $this->assertEquals($expectedVal, $object->getInitials());
    }

    public function testGetSetName()
    {
        $expectedVal = 'AABBCCDD';

        $object = new Author();
        $this->assertNull($object->getName());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment\\Author', $object->setName($expectedVal));
        $this->assertEquals($expectedVal, $object->getName());
    }
}
