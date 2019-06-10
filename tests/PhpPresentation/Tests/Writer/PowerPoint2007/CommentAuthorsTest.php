<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class CommentAuthorsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testComments()
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

    public function testWithoutComment()
    {
        $this->assertZipFileNotExists('ppt/commentAuthors.xml');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testWithoutCommentAuthor()
    {
        $oComment = new Comment();
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $this->assertZipFileNotExists('ppt/commentAuthors.xml');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testWithSameAuthor()
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
