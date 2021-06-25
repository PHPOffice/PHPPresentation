<?php

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Slide\Background\Image;
use PHPUnit\Framework\TestCase;

class ImageTest extends TestCase
{
    public function testColor(): void
    {
        $object = new Image();

        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';
        $numSlide = (string) mt_rand(1, 100);

        $this->assertNull($object->getPath());
        $this->assertEmpty($object->getFilename());
        $this->assertEmpty($object->getExtension());
        $this->assertEquals('background_' . $numSlide . '.', $object->getIndexedFilename($numSlide));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Image', $object->setPath($imagePath));
        $this->assertEquals($imagePath, $object->getPath());
        $this->assertEquals('PhpPresentationLogo.png', $object->getFilename());
        $this->assertEquals('png', $object->getExtension());
        $this->assertEquals('background_' . $numSlide . '.png', $object->getIndexedFilename($numSlide));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Image', $object->setPath('', false));
        $this->assertEquals('', $object->getPath());
        $this->assertEmpty($object->getFilename());
        $this->assertEmpty($object->getExtension());
        $this->assertEquals('background_' . $numSlide . '.', $object->getIndexedFilename($numSlide));
    }

    public function testPathException(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('File not found :');

        $object = new Image();
        $object->setPath('pathDoesntExist', true);
    }
}
