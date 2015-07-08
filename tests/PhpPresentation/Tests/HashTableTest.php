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

use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\Slide;

/**
 * Test class for HashTable
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\HashTable
 */
class HashTableTest extends \PHPUnit_Framework_TestCase
{
    /**
     */
    public function testConstructNull()
    {
        $object = new HashTable();

        $this->assertEquals(0, $object->count());
        $this->assertNull($object->getByIndex());
        $this->assertNull($object->getByHashCode());
        $this->assertInternalType('array', $object->toArray());
        $this->assertEmpty($object->toArray());
    }

    /**
     */
    public function testConstructSource()
    {
        $object = new HashTable(array(
            new Slide(),
            new Slide(),
        ));

        $this->assertEquals(2, $object->count());
        $this->assertInternalType('array', $object->toArray());
        $this->assertCount(2, $object->toArray());
    }

    /**
     */
    public function testAdd()
    {
        $object = new HashTable();
        $oSlide = new Slide();

        // Add From Source : Null
        $this->assertNull($object->addFromSource());
        // Add From Source : Array
        $this->assertNull($object->addFromSource(array($oSlide)));
        $this->assertInternalType('array', $object->toArray());
        $this->assertCount(1, $object->toArray());
        // Clear
        $this->assertNull($object->clear());
        $this->assertEmpty($object->toArray());
        // Add Object
        $this->assertNull($object->add($oSlide));
        $this->assertCount(1, $object->toArray());
        $this->assertNull($object->clear());
        // Add Object w/Hash Index
        $oSlide->setHashIndex(rand(1, 100));
        $this->assertNull($object->add($oSlide));
        $this->assertCount(1, $object->toArray());
        // Add Object w/the same Hash Index
        $this->assertNull($object->add($oSlide));
        $this->assertCount(1, $object->toArray());
    }

    /**
     */
    public function testIndex()
    {
        $object = new HashTable();
        $oSlide1 = new Slide();
        $oSlide2 = new Slide();

        // Add Object
        $this->assertNull($object->add($oSlide1));
        $this->assertNull($object->add($oSlide2));
        // Index
        $this->assertEquals(0, $object->getIndexForHashCode($oSlide1->getHashCode()));
        $this->assertEquals(1, $object->getIndexForHashCode($oSlide2->getHashCode()));
        $this->assertEquals($oSlide1, $object->getByIndex(0));
        $this->assertEquals($oSlide2, $object->getByIndex(1));
    }

    /**
     */
    public function testRemove()
    {
        $object = new HashTable();
        $oSlide1 = new Slide();
        $oSlide2 = new Slide();
        $oSlide3 = new Slide();

        // Add Object
        $this->assertNull($object->add($oSlide1));
        $this->assertNull($object->add($oSlide2));
        $this->assertNull($object->add($oSlide3));
        // Remove
        $this->assertNull($object->remove($oSlide2));
        $this->assertCount(2, $object->toArray());
        $this->assertNull($object->remove($oSlide3));
        $this->assertCount(1, $object->toArray());
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Invalid array parameter passed.
     */
    public function testAddException()
    {
        $object = new HashTable();
        $oSlide = new Slide();
        $object->addFromSource($oSlide);
    }
}
