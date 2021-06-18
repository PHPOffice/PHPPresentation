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
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Writer\Serialized;
use PHPUnit\Framework\TestCase;

/**
 * Test class for serialized reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\Serialized
 */
class SerializedTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Serialized(new PhpPresentation());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
    }

    public function testEmptyConstruct(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('No PhpPresentation assigned.');

        $object = new Serialized();
        $object->getPhpPresentation();
    }

    public function testSaveEmpty(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('Filename is empty.');

        $object = new Serialized(new PhpPresentation());
        $object->save('');
    }

    public function testSaveNoObject(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('No PhpPresentation assigned.');

        $object = new Serialized();
        $object->save('file.phpppt');
    }

    public function testSave(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png');
        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');
        $object = new Serialized($oPhpPresentation);
        $object->save($file);

        $this->assertFileExists($file);
    }

    public function testSaveNotExistingDir(): void
    {
        $this->expectException(\Exception::class);
        $this->expectExceptionMessage('Could not open');

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png');
        $object = new Serialized($oPhpPresentation);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');

        $object->save($file . DIRECTORY_SEPARATOR . 'test' . DIRECTORY_SEPARATOR . 'test');
    }

    public function testSaveOverwriting(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');
        file_put_contents($file, rand(1, 100));

        $object = new Serialized($oPhpPresentation);
        $object->save($file);

        $this->assertFileExists($file);
    }
}
