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
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide\Note;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class NoteTest extends TestCase
{
    public function testParent(): void
    {
        $object = new Note();
        self::assertNull($object->getParent());

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->createSlide();
        $oSlide->setNote($object);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide', $object->getParent());
    }

    public function testExtent(): void
    {
        $object = new Note();
        self::assertNotNull($object->getExtentX());

        $object = new Note();
        self::assertNotNull($object->getExtentY());
    }

    public function testHashCode(): void
    {
        $object = new Note();
        self::assertIsString($object->getHashCode());
    }

    public function testOffset(): void
    {
        $object = new Note();
        self::assertNotNull($object->getOffsetX());

        $object = new Note();
        self::assertNotNull($object->getOffsetY());
    }

    public function testShape(): void
    {
        $object = new Note();
        self::assertCount(0, $object->getShapeCollection());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $object->createRichTextShape());
        self::assertCount(1, $object->getShapeCollection());

        $oRichText = new RichText();
        self::assertInstanceOf(Note::class, $object->addShape($oRichText));
        self::assertCount(2, $object->getShapeCollection());
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $object = new Note();
        $value = mt_rand(1, 100);
        self::assertNull($object->getHashIndex());
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }
}
