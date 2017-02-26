<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class CommentAuthorsTest extends PhpPresentationTestCase
{
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

        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml');
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml', $expectedElement, 'id', 0);
        $this->assertZipXmlAttributeEquals($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml', $expectedElement, 'name', $expectedName);
        $this->assertZipXmlAttributeEquals($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml', $expectedElement, 'initials', $expectedInitials);
    }

    public function testWithoutComment()
    {
        $this->assertZipFileNotExists($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml');
    }

    public function testWithoutCommentAuthor()
    {
        $oComment = new Comment();
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $this->assertZipFileNotExists($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml');
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

        $this->assertZipFileExists($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml');
        $this->assertZipXmlElementExists($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml', $expectedElement);
        $this->assertZipXmlElementCount($this->oPresentation, 'PowerPoint2007', 'ppt/commentAuthors.xml', $expectedElement, 1);
    }
}
