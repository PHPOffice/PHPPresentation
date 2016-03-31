<?php

namespace PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\Media;

class MediaTest extends \PHPUnit_Framework_TestCase
{
    public function testInheritance()
    {
        $object = new Media();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing', $object);
    }
}
