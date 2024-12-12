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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Comment\Author;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Chart element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Comment
 */
class CommentTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Comment();

        self::assertNull($object->getAuthor());
        self::assertNull($object->getText());
        self::assertIsInt($object->getDate());
        self::assertNull($object->getHeight());
        self::assertNull($object->getWidth());
    }

    public function testGetSetAuthor(): void
    {
        $object = new Comment();

        /** @var Author $oStub */
        $oStub = $this->getMockBuilder(Author::class)->getMock();

        self::assertNull($object->getAuthor());
        self::assertInstanceOf(Comment::class, $object->setAuthor($oStub));
        self::assertInstanceOf(Author::class, $object->getAuthor());
    }

    public function testGetSetDate(): void
    {
        $expectedDate = time();

        $object = new Comment();
        self::assertIsInt($object->getDate());
        self::assertInstanceOf(Comment::class, $object->setDate($expectedDate));
        self::assertEquals($expectedDate, $object->getDate());
        self::assertIsInt($object->getDate());
    }

    public function testGetSetText(): void
    {
        $expectedText = 'AABBCCDD';

        $object = new Comment();
        self::assertNull($object->getText());
        self::assertInstanceOf(Comment::class, $object->setText($expectedText));
        self::assertEquals($expectedText, $object->getText());
    }

    public function testGetSetHeightAndWidtg(): void
    {
        $object = new Comment();
        self::assertNull($object->getHeight());
        self::assertNull($object->getWidth());
        self::assertInstanceOf(Comment::class, $object->setHeight(1));
        self::assertInstanceOf(Comment::class, $object->setWidth(1));
        self::assertNull($object->getHeight());
        self::assertNull($object->getWidth());
    }
}
