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

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide\Note;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PhpPresentation
 */
class NoteTest extends TestCase
{
    public function testParent()
    {
        $object = new Note();
        static::assertNull($object->getParent());

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->createSlide();
        $oSlide->setNote($object);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide', $object->getParent());
    }

    public function testExtent()
    {
        $object = new Note();
        static::assertNotNull($object->getExtentX());

        $object = new Note();
        static::assertNotNull($object->getExtentY());
    }

    public function testHashCode()
    {
        $object = new Note();
        static::assertInternalType('string', $object->getHashCode());
    }

    public function testOffset()
    {
        $object = new Note();
        static::assertNotNull($object->getOffsetX());

        $object = new Note();
        static::assertNotNull($object->getOffsetY());
    }

    public function testShape()
    {
        $object = new Note();
        static::assertEquals(0, $object->getShapeCollection()->count());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $object->createRichTextShape());
        static::assertEquals(1, $object->getShapeCollection()->count());

        $oRichText = new RichText();
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $object->addShape($oRichText));
        static::assertEquals(2, $object->getShapeCollection()->count());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Note();
        $value = mt_rand(1, 100);
        static::assertNull($object->getHashIndex());
        $object->setHashIndex($value);
        static::assertEquals($value, $object->getHashIndex());
    }
}
