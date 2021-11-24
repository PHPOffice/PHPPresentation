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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\Exception\DirectoryNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
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

    public function testSaveEmpty(): void
    {
        $this->expectException(InvalidParameterException::class);
        $this->expectExceptionMessage('The parameter pFilename can\'t have the value ""');

        $object = new Serialized(new PhpPresentation());
        $object->save('');
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
        $path = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized' . DIRECTORY_SEPARATOR . 'test');

        $this->expectException(DirectoryNotFoundException::class);
        $this->expectExceptionMessage(sprintf(
            'The directory %s doesn\'t exist',
            $path
        ));

        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oImage = $oSlide->createDrawingShape();
        $oImage->setPath(PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png');

        $object = new Serialized($oPhpPresentation);
        $object->save($path . DIRECTORY_SEPARATOR . 'test');
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
