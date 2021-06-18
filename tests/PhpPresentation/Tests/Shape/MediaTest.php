<?php

namespace PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\Media;
use PHPUnit\Framework\TestCase;

class MediaTest extends TestCase
{
    public function testInheritance(): void
    {
        $object = new Media();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\File', $object);
    }

    public function testMimeType(): void
    {
        $object = new Media();
        $object->setPath('file.mp4', false);
        $this->assertEquals('video/mp4', $object->getMimeType());
        $object->setPath('file.ogv', false);
        $this->assertEquals('video/ogg', $object->getMimeType());
        $object->setPath('file.wmv', false);
        $this->assertEquals('video/x-ms-wmv', $object->getMimeType());
        $object->setPath('file.xxx', false);
        $this->assertEquals('application/octet-stream', $object->getMimeType());
    }
}
