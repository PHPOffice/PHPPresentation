<?php
/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\MemoryDrawing;

/**
 * Test class for memory drawing element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\MemoryDrawing
 */
class MemoryDrawingTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testConstruct()
    {
        $object = new MemoryDrawing();

        $this->assertEquals('imagepng', $object->getRenderingFunction());
        $this->assertEquals(MemoryDrawing::MIMETYPE_DEFAULT, $object->getMimeType());
        $this->assertNull($object->getImageResource());
        $this->assertInternalType('string', $object->getIndexedFilename());
        $this->assertInternalType('string', $object->getExtension());
        $this->assertInternalType('string', $object->getHashCode());
    }

    public function testImageResource()
    {
        $object = new MemoryDrawing();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\MemoryDrawing', $object->setImageResource());
        $this->assertNull($object->getImageResource());
        $this->assertEquals(0, $object->getWidth());
        $this->assertEquals(0, $object->getHeight());

        $width = rand(1, 100);
        $height = rand(100, 200);
        $gdImage = @imagecreatetruecolor($width, $height);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\MemoryDrawing', $object->setImageResource($gdImage));
        $this->assertTrue(is_resource($object->getImageResource()));
        $this->assertEquals($width, $object->getWidth());
        $this->assertEquals($height, $object->getHeight());
    }

    public function testMimeType()
    {
        $object = new MemoryDrawing();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\MemoryDrawing', $object->setMimeType());
        $this->assertEquals(MemoryDrawing::MIMETYPE_DEFAULT, $object->getMimeType());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\MemoryDrawing', $object->setMimeType(MemoryDrawing::MIMETYPE_JPEG));
        $this->assertEquals(MemoryDrawing::MIMETYPE_JPEG, $object->getMimeType());
    }

    public function testRenderingFunction()
    {
        $object = new MemoryDrawing();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\MemoryDrawing', $object->setRenderingFunction());
        $this->assertEquals(MemoryDrawing::RENDERING_DEFAULT, $object->getRenderingFunction());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\MemoryDrawing', $object->setRenderingFunction(MemoryDrawing::RENDERING_JPEG));
        $this->assertEquals(MemoryDrawing::RENDERING_JPEG, $object->getRenderingFunction());
    }
}
