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

        $this->assertNull($object->getAuthor());
        $this->assertNull($object->getText());
        $this->assertIsInt($object->getDate());
        $this->assertNull($object->getHeight());
        $this->assertNull($object->getWidth());
    }

    public function testGetSetAuthor(): void
    {
        $object = new Comment();

        /** @var Author $oStub */
        $oStub = $this->getMockBuilder(Author::class)->getMock();

        $this->assertNull($object->getAuthor());
        $this->assertInstanceOf(Comment::class, $object->setAuthor($oStub));
        $this->assertInstanceOf(Author::class, $object->getAuthor());
    }

    public function testGetSetDate(): void
    {
        $expectedDate = time();

        $object = new Comment();
        $this->assertIsInt($object->getDate());
        $this->assertInstanceOf(Comment::class, $object->setDate($expectedDate));
        $this->assertEquals($expectedDate, $object->getDate());
        $this->assertIsInt($object->getDate());
    }

    public function testGetSetText(): void
    {
        $expectedText = 'AABBCCDD';

        $object = new Comment();
        $this->assertNull($object->getText());
        $this->assertInstanceOf(Comment::class, $object->setText($expectedText));
        $this->assertEquals($expectedText, $object->getText());
    }

    public function testGetSetHeightAndWidtg(): void
    {
        $object = new Comment();
        $this->assertNull($object->getHeight());
        $this->assertNull($object->getWidth());
        $this->assertInstanceOf(Comment::class, $object->setHeight(1));
        $this->assertInstanceOf(Comment::class, $object->setWidth(1));
        $this->assertNull($object->getHeight());
        $this->assertNull($object->getWidth());
    }
}
