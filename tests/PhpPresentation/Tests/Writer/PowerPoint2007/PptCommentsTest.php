<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptCommentsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testComments()
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
    public function testWithoutComment()
    {
        $this->assertZipFileNotExists('ppt/comments/comment1.xml');
        $this->assertIsSchemaECMA376Valid();
    }
}
