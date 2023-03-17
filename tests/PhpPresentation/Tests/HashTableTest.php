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

use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\Slide;
use PHPUnit\Framework\TestCase;

/**
 * Test class for HashTable.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\HashTable
 */
class HashTableTest extends TestCase
{
    public function testConstructNull(): void
    {
        $object = new HashTable();

        $this->assertEquals(0, $object->count());
        $this->assertNull($object->getByIndex());
        $this->assertNull($object->getByHashCode());
        $this->assertIsArray($object->toArray());
        $this->assertEmpty($object->toArray());
    }

    public function testConstructSource(): void
    {
        $object = new HashTable([
            new Slide(),
            new Slide(),
        ]);

        $this->assertEquals(2, $object->count());
        $this->assertIsArray($object->toArray());
        $this->assertCount(2, $object->toArray());
    }

    public function testAdd(): void
    {
        $object = new HashTable();
        $oSlide = new Slide();

        // Add From Source : Null
        $object->addFromSource();
        // Add From Source : Array
        $object->addFromSource([$oSlide]);
        $this->assertIsArray($object->toArray());
        $this->assertCount(1, $object->toArray());
        // Clear
        $object->clear();
        $this->assertEmpty($object->toArray());
        // Add Object
        $object->add($oSlide);
        $this->assertCount(1, $object->toArray());
        $object->clear();
        // Add Object w/Hash Index
        $oSlide->setHashIndex(rand(1, 100));
        $object->add($oSlide);
        $this->assertCount(1, $object->toArray());
        // Add Object w/the same Hash Index
        $object->add($oSlide);
        $this->assertCount(1, $object->toArray());
    }

    public function testIndex(): void
    {
        $object = new HashTable();
        $oSlide1 = new Slide();
        $oSlide2 = new Slide();

        // Add Object
        $object->add($oSlide1);
        $object->add($oSlide2);
        // Index
        $this->assertEquals(0, $object->getIndexForHashCode($oSlide1->getHashCode()));
        $this->assertEquals(1, $object->getIndexForHashCode($oSlide2->getHashCode()));
        $this->assertEquals($oSlide1, $object->getByIndex(0));
        $this->assertEquals($oSlide2, $object->getByIndex(1));
    }

    public function testRemove(): void
    {
        $object = new HashTable();
        $oSlide1 = new Slide();
        $oSlide2 = new Slide();
        $oSlide3 = new Slide();

        // Add Object
        $object->add($oSlide1);
        $object->add($oSlide2);
        $object->add($oSlide3);
        // Remove
        $object->remove($oSlide2);
        $this->assertCount(2, $object->toArray());
        $object->remove($oSlide3);
        $this->assertCount(1, $object->toArray());
    }
}
