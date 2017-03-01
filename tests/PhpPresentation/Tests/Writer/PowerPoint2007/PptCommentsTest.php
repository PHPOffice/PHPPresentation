<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptCommentsTest extends PhpPresentationTestCase
{
    public function testComments()
    {
        $expectedElement = '/p:cmLst/p:cm';

        $oAuthor = new Comment\Author();
        $oComment = new Comment();
        $oComment->setAuthor($oAuthor);
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'ppt/comments/comment1.xml');
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'ppt/comments/comment1.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals($this->oPresentation, 'PowerPoint2007', 'ppt/comments/comment1.xml', $expectedElement, 'authorId', 0);
    }
    public function testWithoutComment()
    {
        $this->assertZipFileNotExists($this->oPresentation, 'PowerPoint2007', 'ppt/comments/comment1.xml');
    }
}
