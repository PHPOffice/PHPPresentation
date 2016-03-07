<?php

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

class PptCommentsTest extends \PHPUnit_Framework_TestCase
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
        $expectedElement = '/p:cmLst/p:cm';

        $oAuthor = new Comment\Author();
        $oComment = new Comment();
        $oComment->setAuthor($oAuthor);
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $pres = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->fileExists('ppt/comments/comment1.xml'));
        $this->assertTrue($pres->elementExists($expectedElement, 'ppt/comments/comment1.xml'));
        $this->assertEquals(0, $pres->getElementAttribute($expectedElement, 'authorId', 'ppt/comments/comment1.xml'));
    }
    public function testWithoutComment()
    {
        $pres = TestHelperDOCX::getDocument($this->oPresentation, 'PowerPoint2007');
        $this->assertFalse($pres->fileExists('ppt/comments/comment1.xml'));
    }
}
