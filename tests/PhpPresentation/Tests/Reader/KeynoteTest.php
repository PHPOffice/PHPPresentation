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
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Reader;

use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\Reader\Keynote;
use PhpOffice\PhpPresentation\Shape\Drawing\Base64;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide;
use PHPUnit\Framework\TestCase;
use ZipArchive;

/**
 * Test class for the Keynote reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\Keynote
 */
class KeynoteTest extends TestCase
{
    /**
     * The `.key` package of a Keynote '13 (and later) presentation : `Index/*.iwa` components.
     *
     * @var string
     */
    protected const FILE_IWA = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Keynote_13.key';

    /**
     * The `.key` package of a Keynote '09 presentation : an `index.apxl` document.
     *
     * @var string
     */
    protected const FILE_APXL = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Keynote_09.key';

    /**
     * @var array<int, string>
     */
    protected $arrayTempFiles = [];

    protected function tearDown(): void
    {
        foreach ($this->arrayTempFiles as $file) {
            if (file_exists($file)) {
                unlink($file);
            }
        }
        $this->arrayTempFiles = [];
    }

    public function testCanRead(): void
    {
        $object = new Keynote();

        self::assertFalse($object->canRead(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt'));
        self::assertFalse($object->canRead(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt'));
        self::assertFalse($object->canRead(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx'));
        self::assertFalse($object->canRead(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.odp'));
        self::assertTrue($object->canRead(self::FILE_IWA));
        self::assertTrue($object->canRead(self::FILE_APXL));
    }

    public function testLoadFileNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new Keynote();
        $object->load('');
    }

    public function testLoadFileBadFormat(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';

        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage(sprintf(
            'The file %s is not in the format supported by class PhpOffice\PhpPresentation\Reader\Keynote',
            $file
        ));

        $object = new Keynote();
        $object->load($file);
    }

    /**
     * A `.key` which is a Zip archive, but holds neither shape of the format, names what is missing
     * rather than being read as an empty presentation.
     */
    public function testLoadPackageWithoutSlide(): void
    {
        $file = $this->createPackage([
            'Index/Document.iwa' => 'not a component',
            'Metadata/DocumentIdentifier' => 'keynote',
        ]);

        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('no slide was found, neither as index.apxl nor as an Index/Slide*.iwa component');

        $object = new Keynote();
        $object->load($file);
    }

    /**
     * A component whose frames are not Snappy compressed at all is refused, rather than being read
     * as text nobody wrote.
     */
    public function testLoadIwaBadFrame(): void
    {
        $file = $this->createPackage([
            'Index/Slide-1.iwa' => "\x42\x01\x00\x00\x00",
        ]);

        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('unsupported Snappy frame type 0x42');

        $object = new Keynote();
        $object->load($file);
    }

    public function testLoadIwa(): void
    {
        $object = new Keynote();
        $oPhpPresentation = $object->load(self::FILE_IWA);

        self::assertEquals(2, $oPhpPresentation->getSlideCount());

        $oSlide = $oPhpPresentation->getSlide(0);
        self::assertEquals(
            ['Title of the first slide', 'A second paragraph'],
            $this->getParagraphs($oSlide)
        );
        self::assertEquals(['The note of the first slide'], $this->getNotes($oSlide));
        self::assertEquals(['PhpPresentationLogo.png'], $this->getImageNames($oSlide));

        $oSlide = $oPhpPresentation->getSlide(1);
        self::assertEquals(['Title of the second slide'], $this->getParagraphs($oSlide));
        self::assertEquals([], $this->getNotes($oSlide));
        self::assertEquals(['photo.jpg'], $this->getImageNames($oSlide));
    }

    /**
     * The text of a master slide is not the text of a slide, and no `MasterSlide` component is read.
     */
    public function testLoadIwaIgnoresMasterSlide(): void
    {
        $object = new Keynote();
        $oPhpPresentation = $object->load(self::FILE_IWA);

        foreach ($oPhpPresentation->getAllSlides() as $oSlide) {
            self::assertNotContains('A master slide nobody reads', $this->getParagraphs($oSlide));
        }
    }

    public function testLoadIwaWithoutImages(): void
    {
        $object = new Keynote();
        $oPhpPresentation = $object->load(self::FILE_IWA, Keynote::SKIP_IMAGES);

        self::assertEquals(2, $oPhpPresentation->getSlideCount());
        foreach ($oPhpPresentation->getAllSlides() as $oSlide) {
            self::assertEquals([], $this->getImageNames($oSlide));
        }
        self::assertEquals(
            ['Title of the first slide', 'A second paragraph'],
            $this->getParagraphs($oPhpPresentation->getSlide(0))
        );
    }

    public function testLoadApxl(): void
    {
        $object = new Keynote();
        $oPhpPresentation = $object->load(self::FILE_APXL);

        self::assertEquals(2, $oPhpPresentation->getSlideCount());

        $oSlide = $oPhpPresentation->getSlide(0);
        self::assertEquals(
            ['Title of the first slide', 'A first paragraph', 'A second paragraph'],
            $this->getParagraphs($oSlide)
        );
        self::assertEquals(['The note of the first slide'], $this->getNotes($oSlide));
        self::assertEquals(['PhpPresentationLogo.png'], $this->getImageNames($oSlide));

        // an empty placeholder is not a text shape, and the thumbnail of the package is not an image
        $oSlide = $oPhpPresentation->getSlide(1);
        self::assertEquals(['Title of the second slide'], $this->getParagraphs($oSlide));
        self::assertEquals([], $this->getImageNames($oSlide));
    }

    /**
     * The position and the size a drawable states are read, so that a slide read out of a `.key`
     * can be written back to another format.
     */
    public function testLoadApxlGeometry(): void
    {
        $object = new Keynote();
        $oSlide = $object->load(self::FILE_APXL)->getSlide(0);

        $shapes = array_values((array) $oSlide->getShapeCollection());

        $oShape = $shapes[0];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(100, $oShape->getOffsetX());
        self::assertEquals(50, $oShape->getOffsetY());
        self::assertEquals(800, $oShape->getWidth());
        self::assertEquals(120, $oShape->getHeight());

        $oShape = $shapes[2];
        self::assertInstanceOf(Base64::class, $oShape);
        self::assertEquals(200, $oShape->getOffsetX());
        self::assertEquals(300, $oShape->getOffsetY());
        self::assertEquals(320, $oShape->getWidth());
        self::assertEquals(240, $oShape->getHeight());
        self::assertStringStartsWith('data:image/png;base64,', $oShape->getData());
    }

    public function testLoadApxlWithoutImages(): void
    {
        $object = new Keynote();
        $oSlide = $object->load(self::FILE_APXL, Keynote::SKIP_IMAGES)->getSlide(0);

        self::assertEquals([], $this->getImageNames($oSlide));
        self::assertEquals(
            ['Title of the first slide', 'A first paragraph', 'A second paragraph'],
            $this->getParagraphs($oSlide)
        );
    }

    /**
     * A `index.apxl.gz`, which is how Keynote '09 compresses the same document, is read too.
     */
    public function testLoadApxlGzipped(): void
    {
        $oZip = new ZipArchive();
        $oZip->open(self::FILE_APXL);
        $contents = $oZip->getFromName('index.apxl');
        $oZip->close();

        self::assertIsString($contents);
        $file = $this->createPackage(['index.apxl.gz' => (string) gzencode($contents)]);

        $object = new Keynote();
        $oPhpPresentation = $object->load($file);

        self::assertEquals(2, $oPhpPresentation->getSlideCount());
        self::assertEquals(
            ['Title of the first slide', 'A first paragraph', 'A second paragraph'],
            $this->getParagraphs($oPhpPresentation->getSlide(0))
        );
    }

    public function testLoadApxlNotGzipped(): void
    {
        $file = $this->createPackage(['index.apxl.gz' => '<key:presentation/>']);

        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('index.apxl.gz is not gzip compressed');

        $object = new Keynote();
        $object->load($file);
    }

    /**
     * A `.key` package holding the given entries, removed once the test is over.
     *
     * @param array<string, string> $arrayEntries
     */
    protected function createPackage(array $arrayEntries): string
    {
        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        self::assertIsString($file);
        $this->arrayTempFiles[] = $file;

        $oZip = new ZipArchive();
        $oZip->open($file, ZipArchive::OVERWRITE);
        foreach ($arrayEntries as $name => $contents) {
            $oZip->addFromString($name, $contents);
        }
        $oZip->close();

        return $file;
    }

    /**
     * Every paragraph of every text shape of a slide.
     *
     * @return array<int, string>
     */
    protected function getParagraphs(Slide $oSlide): array
    {
        $paragraphs = [];
        foreach ($oSlide->getShapeCollection() as $oShape) {
            if ($oShape instanceof RichText) {
                foreach ($oShape->getParagraphs() as $oParagraph) {
                    $paragraphs[] = $oParagraph->getPlainText();
                }
            }
        }

        return $paragraphs;
    }

    /**
     * Every paragraph of the speaker note of a slide.
     *
     * @return array<int, string>
     */
    protected function getNotes(Slide $oSlide): array
    {
        $paragraphs = [];
        foreach ($oSlide->getNote()->getShapeCollection() as $oShape) {
            if ($oShape instanceof RichText) {
                foreach ($oShape->getParagraphs() as $oParagraph) {
                    $paragraphs[] = $oParagraph->getPlainText();
                }
            }
        }

        return $paragraphs;
    }

    /**
     * The name of every image of a slide.
     *
     * @return array<int, string>
     */
    protected function getImageNames(Slide $oSlide): array
    {
        $names = [];
        foreach ($oSlide->getShapeCollection() as $oShape) {
            if ($oShape instanceof Base64) {
                $names[] = $oShape->getName();
            }
        }

        return $names;
    }
}
