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

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class CommentAuthorsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testComments(): void
    {
        $expectedElement = '/p:cmAuthorLst/p:cmAuthor';
        $expectedName = 'Name';
        $expectedInitials = 'Initials';

        $oAuthor = new Comment\Author();
        $oAuthor->setName($expectedName);
        $oAuthor->setInitials($expectedInitials);
        $oComment = new Comment();
        $oComment->setAuthor($oAuthor);
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $this->assertZipFileExists('ppt/commentAuthors.xml');
        $this->assertZipXmlElementExists('ppt/commentAuthors.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/commentAuthors.xml', $expectedElement, 'id', 0);
        $this->assertZipXmlAttributeEquals('ppt/commentAuthors.xml', $expectedElement, 'name', $expectedName);
        $this->assertZipXmlAttributeEquals('ppt/commentAuthors.xml', $expectedElement, 'initials', $expectedInitials);
        $this->assertIsSchemaECMA376Valid();
    }

    public function testWithoutComment(): void
    {
        $this->assertZipFileNotExists('ppt/commentAuthors.xml');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testWithoutCommentAuthor(): void
    {
        $oComment = new Comment();
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $this->assertZipFileNotExists('ppt/commentAuthors.xml');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testWithSameAuthor(): void
    {
        $expectedElement = '/p:cmAuthorLst/p:cmAuthor';

        $oAuthor = new Comment\Author();

        $oComment1 = new Comment();
        $oComment1->setAuthor($oAuthor);
        $this->oPresentation->getActiveSlide()->addShape($oComment1);
        $oComment2 = new Comment();
        $oComment2->setAuthor($oAuthor);
        $this->oPresentation->getActiveSlide()->addShape($oComment2);

        $this->assertZipFileExists('ppt/commentAuthors.xml');
        $this->assertZipXmlElementExists('ppt/commentAuthors.xml', $expectedElement);
        $this->assertZipXmlElementCount('ppt/commentAuthors.xml', $expectedElement, 1);
        $this->assertIsSchemaECMA376Valid();
    }
}
