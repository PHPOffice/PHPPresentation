<?php

namespace PhpOffice\PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

class RelationshipsTest extends \PHPUnit_Framework_TestCase
{
    public function testCommentsAuthors()
    {
        $oAuthor = new Comment\Author();
        $oComment = new Comment();
        $oComment->setAuthor($oAuthor);
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->addShape($oComment);

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->elementExists('/Relationships/Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"]', 'ppt/_rels/presentation.xml.rels'));
    }
}
