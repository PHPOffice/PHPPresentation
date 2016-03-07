<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

class CommentAuthorsTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var PhpPresentation;
     */
    protected $oPresentation;

    public function setUp()
    {
        $this->oPresentation = new PhpPresentation();
    }

    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        $this->oPresentation = null;
        TestHelperDOCX::clear();
    }

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

        $pres = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->fileExists('ppt/commentAuthors.xml'));
        $this->assertTrue($pres->elementExists($expectedElement, 'ppt/commentAuthors.xml'));
        $this->assertEquals(0, $pres->getElementAttribute($expectedElement, 'id', 'ppt/commentAuthors.xml'));
        $this->assertEquals($expectedName, $pres->getElementAttribute($expectedElement, 'name', 'ppt/commentAuthors.xml'));
        $this->assertEquals($expectedInitials, $pres->getElementAttribute($expectedElement, 'initials', 'ppt/commentAuthors.xml'));
    }

    public function testWithoutComment()
    {
        $pres = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $this->assertFalse($pres->fileExists('ppt/commentAuthors.xml'));
    }

    public function testWithoutCommentAuthor()
    {
        $oComment = new Comment();
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $pres = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $this->assertFalse($pres->fileExists('ppt/commentAuthors.xml'));
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

        $pres = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->fileExists('ppt/commentAuthors.xml'));
        $this->assertTrue($pres->fileExists('ppt/commentAuthors.xml'));
        $this->assertTrue($pres->elementExists($expectedElement, 'ppt/commentAuthors.xml'));
        $this->assertEquals(1, $pres->elementCount($expectedElement, 'ppt/commentAuthors.xml'));
    }
}
