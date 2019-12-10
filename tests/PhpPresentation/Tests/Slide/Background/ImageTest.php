<?php

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Slide\Background\Image;
use PHPUnit\Framework\TestCase;

class ImageTest extends TestCase
{
    public function testColor()
    {
        $object = new Image();

        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';
        $numSlide = mt_rand(1, 100);

        static::assertNull($object->getPath());
        static::assertEmpty($object->getFilename());
        static::assertEmpty($object->getExtension());
        static::assertEquals('background_' . $numSlide . '.', $object->getIndexedFilename($numSlide));

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Image', $object->setPath($imagePath));
        static::assertEquals($imagePath, $object->getPath());
        static::assertEquals('PhpPresentationLogo.png', $object->getFilename());
        static::assertEquals('png', $object->getExtension());
        static::assertEquals('background_' . $numSlide . '.png', $object->getIndexedFilename($numSlide));

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Image', $object->setPath(null, false));
        static::assertNull($object->getPath());
        static::assertEmpty($object->getFilename());
        static::assertEmpty($object->getExtension());
        static::assertEquals('background_' . $numSlide . '.', $object->getIndexedFilename($numSlide));
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage File not found :
     */
    public function testPathException()
    {
        $object = new Image();
        $object->setPath('pathDoesntExist', true);
    }
}
