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

class PptCommentsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testComments(): void
    {
        $expectedElement = '/p:cmLst/p:cm';

        $oAuthor = new Comment\Author();
        $oComment = new Comment();
        $oComment->setAuthor($oAuthor);
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $this->assertZipFileExists('ppt/comments/comment1.xml');
        $this->assertZipXmlElementExists('ppt/comments/comment1.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/comments/comment1.xml', $expectedElement, 'authorId', 0);
        $this->assertIsSchemaECMA376Valid();
    }

    public function testWithoutComment(): void
    {
        $this->assertZipFileNotExists('ppt/comments/comment1.xml');
        $this->assertIsSchemaECMA376Valid();
    }
}
