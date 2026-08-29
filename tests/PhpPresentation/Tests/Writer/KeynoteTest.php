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

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\Exception\DirectoryNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Reader\Keynote as KeynoteReader;
use PhpOffice\PhpPresentation\Shape\Drawing\Base64;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\Keynote;
use ZipArchive;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\Keynote.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\Keynote
 */
class KeynoteTest extends PhpPresentationTestCase
{
    protected $writerName = 'Keynote';

    public function testConstruct(): void
    {
        $object = new Keynote($this->oPresentation);

        self::assertInstanceOf(PhpPresentation::class, $object->getPhpPresentation());
        self::assertInstanceOf(HashTable::class, $object->getDrawingHashTable());
    }

    public function testConstructWithoutPresentation(): void
    {
        $object = new Keynote();

        self::assertInstanceOf(PhpPresentation::class, $object->getPhpPresentation());
    }

    public function testSaveEmptyFilename(): void
    {
        $this->expectException(InvalidParameterException::class);

        $object = new Keynote($this->oPresentation);
        $object->save('');
    }

    public function testSaveDirectoryNotExists(): void
    {
        $this->expectException(DirectoryNotFoundException::class);

        $object = new Keynote($this->oPresentation);
        $object->save('/directory/which/does/not/exist/sample.key');
    }

    /**
     * The package holds an `index.apxl` document and one file per image of the presentation.
     */
    public function testSave(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->createRichTextShape()->createTextRun('A text run');
        $oSlide->createDrawingShape()
            ->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');

        $object = new Keynote($this->oPresentation);
        $object->save($this->filePath);

        $oZip = new ZipArchive();
        self::assertTrue($oZip->open($this->filePath));

        $entries = [];
        for ($index = 0; $index < $oZip->numFiles; ++$index) {
            $entries[] = (string) $oZip->getNameIndex($index);
        }
        $oZip->close();

        self::assertContains('index.apxl', $entries);
        self::assertCount(2, $entries);
        self::assertMatchesRegularExpression('#^Data/PhpPresentationLogo[0-9]+\.png$#', $entries[1]);
    }

    /**
     * What the writer writes, the reader reads : the text of every slide, its speaker note, the
     * images it uses and where they sit.
     */
    public function testSaveAndLoad(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->setOffsetX(10)->setOffsetY(20)->setWidth(300)->setHeight(100);
        $oShape->createTextRun('The first paragraph');
        $oShape->createParagraph()->createTextRun('The second paragraph');
        $oSlide->createDrawingShape()
            ->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png')
            ->setOffsetX(30)
            ->setOffsetY(40);
        $oSlide->getNote()->createRichTextShape()->createTextRun('A speaker note');

        $this->oPresentation->createSlide()->createRichTextShape()->createTextRun('The second slide');

        $oPhpPresentation = $this->writeAndLoad();

        self::assertEquals(2, $oPhpPresentation->getSlideCount());

        $shapes = array_values((array) $oPhpPresentation->getSlide(0)->getShapeCollection());
        self::assertCount(2, $shapes);

        self::assertInstanceOf(RichText::class, $shapes[0]);
        self::assertEquals('The first paragraphThe second paragraph', $shapes[0]->getPlainText());
        self::assertEquals(10, $shapes[0]->getOffsetX());
        self::assertEquals(20, $shapes[0]->getOffsetY());
        self::assertEquals(300, $shapes[0]->getWidth());
        self::assertEquals(100, $shapes[0]->getHeight());

        self::assertInstanceOf(Base64::class, $shapes[1]);
        self::assertMatchesRegularExpression('#^PhpPresentationLogo[0-9]+\.png$#', $shapes[1]->getName());
        self::assertEquals(30, $shapes[1]->getOffsetX());
        self::assertEquals(40, $shapes[1]->getOffsetY());

        $notes = array_values((array) $oPhpPresentation->getSlide(0)->getNote()->getShapeCollection());
        self::assertCount(1, $notes);
        self::assertInstanceOf(RichText::class, $notes[0]);
        self::assertEquals('A speaker note', $notes[0]->getPlainText());

        $shapes = array_values((array) $oPhpPresentation->getSlide(1)->getShapeCollection());
        self::assertCount(1, $shapes);
        self::assertInstanceOf(RichText::class, $shapes[0]);
        self::assertEquals('The second slide', $shapes[0]->getPlainText());
    }

    /**
     * A shape holding no text is not written as an empty placeholder, and a slide with no note is
     * written without one.
     */
    public function testSaveEmptyShapes(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->createRichTextShape();
        $oSlide->getNote()->createRichTextShape();

        $oPhpPresentation = $this->writeAndLoad();

        self::assertEquals(1, $oPhpPresentation->getSlideCount());
        self::assertCount(0, $oPhpPresentation->getSlide(0)->getShapeCollection());
        self::assertCount(0, $oPhpPresentation->getSlide(0)->getNote()->getShapeCollection());
    }

    /**
     * Writes the presentation of the test, then reads the package back.
     */
    protected function writeAndLoad(): PhpPresentation
    {
        $object = new Keynote($this->oPresentation);
        $object->save($this->filePath);

        self::assertFileExists($this->filePath);

        return (new KeynoteReader())->load($this->filePath);
    }
}
