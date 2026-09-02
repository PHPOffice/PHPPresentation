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
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Reader\ODPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Field;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide\Background\Color as BackgroundColor;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Writer\ODPresentation as ODPresentationWriter;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use ZipArchive;

/**
 * Test class for ODPresentation reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\ODPresentation
 */
class ODPresentationTest extends TestCase
{
    /**
     * Test can read.
     */
    public function testCanRead(): void
    {
        $object = new ODPresentation();

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt';
        self::assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
        self::assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        self::assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.odp';
        self::assertTrue($object->canRead($file));
    }

    public function testLoadFileNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new ODPresentation();
        $object->load('');
    }

    public function testLoadFileBadFormat(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt';
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage(sprintf(
            'The file %s is not in the format supported by class PhpOffice\PhpPresentation\Reader\ODPresentation',
            $file
        ));

        $object = new ODPresentation();
        $object->load($file);
    }

    public function testFileSupportsNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new ODPresentation();
        $object->fileSupportsUnserializePhpPresentation('');
    }

    public function testLoadFile01(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.odp';
        $object = new ODPresentation();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        // Document Properties
        self::assertEquals('PHPOffice', $oPhpPresentation->getDocumentProperties()->getCreator());
        self::assertEquals('PHPPresentation Team', $oPhpPresentation->getDocumentProperties()->getLastModifiedBy());
        self::assertEquals('Sample 02 Title', $oPhpPresentation->getDocumentProperties()->getTitle());
        self::assertEquals('Sample 02 Subject', $oPhpPresentation->getDocumentProperties()->getSubject());
        self::assertEquals('Sample 02 Description', $oPhpPresentation->getDocumentProperties()->getDescription());
        self::assertEquals('office 2007 openxml libreoffice odt php', $oPhpPresentation->getDocumentProperties()->getKeywords());
        self::assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        // Presentation Properties
        self::assertEquals(PresentationProperties::SLIDESHOW_TYPE_PRESENT, $oPhpPresentation->getPresentationProperties()->getSlideshowType());

        self::assertCount(4, $oPhpPresentation->getAllSlides());

        // Slide 1
        $oSlide1 = $oPhpPresentation->getSlide(0);
        $arrayShape = (array) $oSlide1->getShapeCollection();
        self::assertCount(2, $arrayShape);
        // Slide 1 : Shape 1
        /** @var Gd $oShape */
        $oShape = $arrayShape[0];
        self::assertInstanceOf(Gd::class, $oShape);
        self::assertEquals('PHPPresentation logo', $oShape->getName());
        self::assertEquals('PHPPresentation logo', $oShape->getDescription());
        self::assertEquals(36, $oShape->getHeight());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(10, $oShape->getOffsetY());
        self::assertEquals('image/png', $oShape->getMimeType());
        self::assertTrue($oShape->getShadow()->isVisible());
        self::assertEquals(45, $oShape->getShadow()->getDirection());
        self::assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 1 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(200, $oShape->getHeight());
        self::assertEquals(600, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(400, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $oParagraph);
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        self::assertEquals(0, $oParagraph->getSpacingAfter());
        self::assertEquals(0, $oParagraph->getSpacingBefore());
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        self::assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(3, $arrayRichText);
        // Slide 1 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Introduction to', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(28, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 1 : Shape 2 : Paragraph 2
        $oRichText = $arrayRichText[1];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $oRichText);
        // Slide 1 : Shape 2 : Paragraph 3
        $oRichText = $arrayRichText[2];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('PHPPresentation', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(60, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());

        // Slide 2
        $oSlide2 = $oPhpPresentation->getSlide(1);
        $arrayShape = (array) $oSlide2->getShapeCollection();
        self::assertCount(3, $arrayShape);
        // Slide 2 : Shape 1
        /** @var Gd $oShape */
        $oShape = $arrayShape[0];
        self::assertInstanceOf(Gd::class, $oShape);
        self::assertEquals('PHPPresentation logo', $oShape->getName());
        self::assertEquals('PHPPresentation logo', $oShape->getDescription());
        self::assertEquals(36, $oShape->getHeight());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(10, $oShape->getOffsetY());
        self::assertEquals('image/png', $oShape->getMimeType());
        self::assertTrue($oShape->getShadow()->isVisible());
        self::assertEquals(45, $oShape->getShadow()->getDirection());
        self::assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 2 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(100, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        // Slide 2 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('What is PHPPresentation?', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(48, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 2 : Shape 3
        /** @var RichText $oShape */
        $oShape = $arrayShape[2];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(600, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(4, $arrayParagraphs);
        // Slide 2 : Shape 3 : Paragraph 1
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('A class library', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 2 : Shape 3 : Paragraph 2
        $oParagraph = $arrayParagraphs[1];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Written in PHP', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 2 : Shape 3 : Paragraph 3
        $oParagraph = $arrayParagraphs[2];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Representing a presentation', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 2 : Shape 3 : Paragraph 4
        $oParagraph = $arrayParagraphs[3];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Supports writing to different file formats', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());

        // Slide 3
        $oSlide2 = $oPhpPresentation->getSlide(2);
        $arrayShape = (array) $oSlide2->getShapeCollection();
        self::assertCount(3, $arrayShape);
        // Slide 3 : Shape 1
        /** @var Gd $oShape */
        $oShape = $arrayShape[0];
        self::assertInstanceOf(Gd::class, $oShape);
        self::assertEquals('PHPPresentation logo', $oShape->getName());
        self::assertEquals('PHPPresentation logo', $oShape->getDescription());
        self::assertEquals(36, $oShape->getHeight());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(10, $oShape->getOffsetY());
        self::assertEquals('image/png', $oShape->getMimeType());
        self::assertTrue($oShape->getShadow()->isVisible());
        self::assertEquals(45, $oShape->getShadow()->getDirection());
        self::assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 3 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(100, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        self::assertEquals(0, $oParagraph->getSpacingAfter());
        self::assertEquals(0, $oParagraph->getSpacingBefore());
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        self::assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        // Slide 3 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('What\'s the point?', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(48, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[2];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(600, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(8, $arrayParagraphs);
        // Slide 3 : Shape 3 : Paragraph 1
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(0, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Generate slide decks', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 2
        $oParagraph = $arrayParagraphs[1];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Represent business data', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 3
        $oParagraph = $arrayParagraphs[2];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Show a family slide show', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 4
        $oParagraph = $arrayParagraphs[3];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('...', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 5
        $oParagraph = $arrayParagraphs[4];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(0, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Export these to different formats', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 6
        $oParagraph = $arrayParagraphs[5];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('PHPPresentation 2007', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 7
        $oParagraph = $arrayParagraphs[6];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Serialized', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 8
        $oParagraph = $arrayParagraphs[7];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('... (more to come) ...', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());

        // Slide 4
        $oSlide3 = $oPhpPresentation->getSlide(3);
        $arrayShape = (array) $oSlide3->getShapeCollection();
        self::assertCount(3, $arrayShape);
        // Slide 4 : Shape 1
        /** @var Gd $oShape */
        $oShape = $arrayShape[0];
        self::assertInstanceOf(Gd::class, $oShape);
        self::assertEquals('PHPPresentation logo', $oShape->getName());
        self::assertEquals('PHPPresentation logo', $oShape->getDescription());
        self::assertEquals(36, $oShape->getHeight());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(10, $oShape->getOffsetY());
        self::assertEquals('image/png', $oShape->getMimeType());
        self::assertTrue($oShape->getShadow()->isVisible());
        self::assertEquals(45, $oShape->getShadow()->getDirection());
        self::assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 4 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(100, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        self::assertEquals(0, $oParagraph->getSpacingAfter());
        self::assertEquals(0, $oParagraph->getSpacingBefore());
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        self::assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        // Slide 4 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Need more info?', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(48, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 4 : Shape 3
        /** @var RichText $oShape */
        $oShape = $arrayShape[2];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(600, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        self::assertEquals(0, $oParagraph->getSpacingAfter());
        self::assertEquals(0, $oParagraph->getSpacingBefore());
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        self::assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(3, $arrayRichText);
        // Slide 4 : Shape 3 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Check the project site on GitHub:', $oRichText->getText());
        self::assertFalse($oRichText->getFont()->isBold());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 4 : Shape 3 : Paragraph 2
        $oRichText = $arrayRichText[1];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $oRichText);
        // Slide 4 : Shape 3 : Paragraph 3
        /** @var RichText\Run $oRichText */
        $oRichText = $arrayRichText[2];
        self::assertInstanceOf(RichText\Run::class, $oRichText);
        self::assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getText());
        self::assertFalse($oRichText->getFont()->isBold());
        self::assertEquals(32, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        self::assertTrue($oRichText->hasHyperlink());
        self::assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getHyperlink()->getUrl());
        //$this->assertEquals('PHPPresentation', $oRichText->getHyperlink()->getTooltip());
    }

    public function testLoadFileWithoutImages(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.odp';
        $object = new ODPresentation();
        $oPhpPresentation = $object->load($file, ODPresentation::SKIP_IMAGES);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        // Document Properties
        self::assertEquals('PHPOffice', $oPhpPresentation->getDocumentProperties()->getCreator());
        self::assertEquals('PHPPresentation Team', $oPhpPresentation->getDocumentProperties()->getLastModifiedBy());
        self::assertEquals('Sample 02 Title', $oPhpPresentation->getDocumentProperties()->getTitle());
        self::assertEquals('Sample 02 Subject', $oPhpPresentation->getDocumentProperties()->getSubject());
        self::assertEquals('Sample 02 Description', $oPhpPresentation->getDocumentProperties()->getDescription());
        self::assertEquals('office 2007 openxml libreoffice odt php', $oPhpPresentation->getDocumentProperties()->getKeywords());
        self::assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        // Presentation Properties
        self::assertEquals(PresentationProperties::SLIDESHOW_TYPE_PRESENT, $oPhpPresentation->getPresentationProperties()->getSlideshowType());

        self::assertCount(4, $oPhpPresentation->getAllSlides());

        // Slide 1
        $oSlide1 = $oPhpPresentation->getSlide(0);
        $arrayShape = (array) $oSlide1->getShapeCollection();
        self::assertCount(1, $arrayShape);
        // Slide 1 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[0];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(200, $oShape->getHeight());
        self::assertEquals(600, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(400, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $oParagraph);
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        self::assertEquals(0, $oParagraph->getSpacingAfter());
        self::assertEquals(0, $oParagraph->getSpacingBefore());
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        self::assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(3, $arrayRichText);
        // Slide 1 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Introduction to', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(28, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 1 : Shape 2 : Paragraph 2
        $oRichText = $arrayRichText[1];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $oRichText);
        // Slide 1 : Shape 2 : Paragraph 3
        $oRichText = $arrayRichText[2];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('PHPPresentation', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(60, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());

        // Slide 2
        $oSlide2 = $oPhpPresentation->getSlide(1);
        $arrayShape = (array) $oSlide2->getShapeCollection();
        self::assertCount(2, $arrayShape);
        // Slide 2 : Shape 1
        /** @var RichText $oShape */
        $oShape = $arrayShape[0];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(100, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        // Slide 2 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('What is PHPPresentation?', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(48, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 2 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(600, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(4, $arrayParagraphs);
        // Slide 2 : Shape 3 : Paragraph 1
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('A class library', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 2 : Shape 3 : Paragraph 2
        $oParagraph = $arrayParagraphs[1];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Written in PHP', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 2 : Shape 3 : Paragraph 3
        $oParagraph = $arrayParagraphs[2];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Representing a presentation', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 2 : Shape 3 : Paragraph 4
        $oParagraph = $arrayParagraphs[3];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Supports writing to different file formats', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());

        // Slide 3
        $oSlide2 = $oPhpPresentation->getSlide(2);
        $arrayShape = (array) $oSlide2->getShapeCollection();
        self::assertCount(2, $arrayShape);
        // Slide 3 : Shape 1
        /** @var RichText $oShape */
        $oShape = $arrayShape[0];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(100, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        self::assertEquals(0, $oParagraph->getSpacingAfter());
        self::assertEquals(0, $oParagraph->getSpacingBefore());
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        self::assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        // Slide 3 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('What\'s the point?', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(48, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 1
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(600, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(8, $arrayParagraphs);
        // Slide 3 : Shape 3 : Paragraph 1
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(0, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Generate slide decks', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 2
        $oParagraph = $arrayParagraphs[1];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Represent business data', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 3
        $oParagraph = $arrayParagraphs[2];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Show a family slide show', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 4
        $oParagraph = $arrayParagraphs[3];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('...', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 5
        $oParagraph = $arrayParagraphs[4];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(0, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Export these to different formats', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 6
        $oParagraph = $arrayParagraphs[5];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('PHPPresentation 2007', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 7
        $oParagraph = $arrayParagraphs[6];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Serialized', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 3 : Shape 3 : Paragraph 8
        $oParagraph = $arrayParagraphs[7];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('... (more to come) ...', $oRichText->getText());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());

        // Slide 4
        $oSlide3 = $oPhpPresentation->getSlide(3);
        $arrayShape = (array) $oSlide3->getShapeCollection();
        self::assertCount(2, $arrayShape);
        // Slide 4 : Shape 1
        /** @var RichText $oShape */
        $oShape = $arrayShape[0];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(100, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        self::assertEquals(0, $oParagraph->getSpacingAfter());
        self::assertEquals(0, $oParagraph->getSpacingBefore());
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        self::assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayRichText);
        // Slide 4 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Need more info?', $oRichText->getText());
        self::assertTrue($oRichText->getFont()->isBold());
        self::assertEquals(48, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 4 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(600, $oShape->getHeight());
        self::assertEquals(930, $oShape->getWidth());
        self::assertEquals(10, $oShape->getOffsetX());
        self::assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        self::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        self::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
        self::assertEquals(0, $oParagraph->getSpacingAfter());
        self::assertEquals(0, $oParagraph->getSpacingBefore());
        self::assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        self::assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        self::assertCount(3, $arrayRichText);
        // Slide 4 : Shape 3 : Paragraph 1
        $oRichText = $arrayRichText[0];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        self::assertEquals('Check the project site on GitHub:', $oRichText->getText());
        self::assertFalse($oRichText->getFont()->isBold());
        self::assertEquals(36, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        // Slide 4 : Shape 3 : Paragraph 2
        $oRichText = $arrayRichText[1];
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $oRichText);
        // Slide 4 : Shape 3 : Paragraph 3
        /** @var RichText\Run $oRichText */
        $oRichText = $arrayRichText[2];
        self::assertInstanceOf(RichText\Run::class, $oRichText);
        self::assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getText());
        self::assertFalse($oRichText->getFont()->isBold());
        self::assertEquals(32, $oRichText->getFont()->getSize());
        self::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
        self::assertTrue($oRichText->hasHyperlink());
        self::assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getHyperlink()->getUrl());
        //$this->assertEquals('PHPPresentation', $oRichText->getHyperlink()->getTooltip());
    }

    public function testSlideName(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/ODP_Slide_Name.odp';
        $object = new ODPresentation();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        self::assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        self::assertEquals('MaDiapo', $oPhpPresentation->getSlide(0)->getName());
    }

    public function testIssue00141(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Issue_00141.odp';
        $object = new ODPresentation();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        self::assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        self::assertCount(3, $oPhpPresentation->getAllSlides());

        // Slide 1
        $oSlide = $oPhpPresentation->getSlide(1);
        $arrayShape = (array) $oSlide->getShapeCollection();
        self::assertCount(2, $arrayShape);
        // Slide 1 : Shape 1
        /** @var RichText $oShape */
        $oShape = reset($arrayShape);
        self::assertInstanceOf(RichText::class, $oShape);
        // Slide 1 : Shape 1 : Paragraph 1
        $oParagraph = $oShape->getParagraph();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $oParagraph);
        // Slide 1 : Shape 1 : Paragraph 1 : RichText Elements
        $arrayElements = $oParagraph->getRichTextElements();
        self::assertCount(1, $arrayElements);
        $oElement = reset($arrayElements);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $oElement);
        self::assertEquals('TEST IMAGE', $oElement->getText());
    }

    /**
     * @return array<array{0: string}>
     */
    public static function dataProviderUnderlines(): array
    {
        return [
            [Font::UNDERLINE_DASH],
            [Font::UNDERLINE_DASHHEAVY],
            [Font::UNDERLINE_DASHLONG],
            [Font::UNDERLINE_DASHLONGHEAVY],
            [Font::UNDERLINE_DOTHASH],
            [Font::UNDERLINE_DOTHASHHEAVY],
            [Font::UNDERLINE_DOTDOTDASH],
            [Font::UNDERLINE_DOTDOTDASHHEAVY],
            [Font::UNDERLINE_DOTTED],
            [Font::UNDERLINE_DOTTEDHEAVY],
            [Font::UNDERLINE_DOUBLE],
            [Font::UNDERLINE_HEAVY],
            [Font::UNDERLINE_SINGLE],
            [Font::UNDERLINE_WAVY],
            [Font::UNDERLINE_WAVYDOUBLE],
            [Font::UNDERLINE_WAVYHEAVY],
            [Font::UNDERLINE_WORDS],
        ];
    }

    /**
     * @dataProvider dataProviderUnderlines
     */
    #[DataProvider('dataProviderUnderlines')]
    public function testFontUnderlineSurvivesTheRoundTrip(string $underline): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oRun = $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('Sample');
        $oRun->getFont()->setUnderline($underline);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        $oFont = $arrayShape[0]->getParagraph()->getRichTextElements()[0]->getFont();
        self::assertInstanceOf(Font::class, $oFont);
        self::assertEquals($underline, $oFont->getUnderline());
    }

    public function testFontStateSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        foreach ([Font::FORMAT_LATIN, Font::FORMAT_EAST_ASIAN, Font::FORMAT_COMPLEX_SCRIPT] as $format) {
            $oRun = $oSlide->createRichTextShape()->createTextRun('Sample');
            $oRun->getFont()
                ->setFormat($format)
                ->setBold(true)
                ->setItalic(true)
                ->setStrikethrough(Font::STRIKE_DOUBLE);
        }
        // a run nobody styled comes back as plain as it went in
        $oSlide->createRichTextShape()->createTextRun('Plain');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertCount(4, $arrayShape);
        foreach ([Font::FORMAT_LATIN, Font::FORMAT_EAST_ASIAN, Font::FORMAT_COMPLEX_SCRIPT] as $idx => $format) {
            self::assertInstanceOf(RichText::class, $arrayShape[$idx]);
            $oFont = $arrayShape[$idx]->getParagraph()->getRichTextElements()[0]->getFont();
            self::assertInstanceOf(Font::class, $oFont);
            self::assertEquals($format, $oFont->getFormat());
            self::assertTrue($oFont->isBold());
            self::assertTrue($oFont->isItalic());
            self::assertEquals(Font::STRIKE_DOUBLE, $oFont->getStrikethrough());
        }
        self::assertInstanceOf(RichText::class, $arrayShape[3]);
        $oFont = $arrayShape[3]->getParagraph()->getRichTextElements()[0]->getFont();
        self::assertInstanceOf(Font::class, $oFont);
        self::assertFalse($oFont->isItalic());
        self::assertEquals(Font::UNDERLINE_NONE, $oFont->getUnderline());
        self::assertEquals(Font::STRIKE_NONE, $oFont->getStrikethrough());
    }

    public function testShapeDecorative(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oSlide->createRichTextShape()->setDecorative()->createTextRun('Decorative');
        $oSlide->createRichTextShape()->createTextRun('Meaningful');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertCount(2, $arrayShape);
        self::assertTrue($arrayShape[0]->isDecorative());
        self::assertFalse($arrayShape[1]->isDecorative());
    }

    public function testShapeDescription(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oRichText = $oSlide->createRichTextShape();
        $oRichText->setDescription('Budget spent to date: 45% of 1.2M EUR');
        $oRichText->createTextRun('45%');
        $oDrawing = $oSlide->createDrawingShape();
        $oDrawing->setName('Logo');
        $oDrawing->setDescription('The logo of the company');
        $oDrawing->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
        // Written by an earlier version: no `svg:desc`, the name is all the shape says.
        $oLegacy = $oSlide->createDrawingShape();
        $oLegacy->setName('Legacy');
        $oLegacy->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertCount(3, $arrayShape);
        self::assertEquals('Budget spent to date: 45% of 1.2M EUR', $arrayShape[0]->getDescription());
        self::assertEquals('The logo of the company', $arrayShape[1]->getDescription());
        self::assertEquals('Legacy', $arrayShape[2]->getDescription());
    }

    public function testHyperlinkToSlide(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oRun = $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('AAAA');
        $oRun->getHyperlink()->setSlideNumber(3);
        $oPhpPresentation->createSlide();
        $oPhpPresentation->createSlide()->setName('Milestone Overview');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        $arrayRichText = $arrayShape[0]->getParagraph()->getRichTextElements();
        self::assertInstanceOf(TextElement::class, $arrayRichText[0]);
        $oHyperlink = $arrayRichText[0]->getHyperlink();

        // Without the lookup this comes back as the literal `#Milestone Overview`
        self::assertTrue($oHyperlink->isInternal());
        self::assertEquals(3, $oHyperlink->getSlideNumber());
    }

    /**
     * @return array<array<string>>
     */
    public static function dataProviderHorizontalAlignment(): array
    {
        return [
            [Alignment::HORIZONTAL_LEFT],
            [Alignment::HORIZONTAL_RIGHT],
            [Alignment::HORIZONTAL_CENTER],
            [Alignment::HORIZONTAL_JUSTIFY],
        ];
    }

    /**
     * @dataProvider dataProviderHorizontalAlignment
     */
    #[DataProvider('dataProviderHorizontalAlignment')]
    public function testHorizontalAlignmentSurvivesTheRoundTrip(string $alignment): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('AAAA');
        $oShape->getActiveParagraph()->getAlignment()->setHorizontal($alignment);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        // and it is the value the model uses, not the one ODF spells it with
        self::assertEquals($alignment, $arrayShape[0]->getParagraph()->getAlignment()->getHorizontal());
    }

    public function testTextDirectionSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('\u{645}\u{631}\u{62d}\u{628}\u{627}');
        $oShape->getActiveParagraph()->getAlignment()->setIsRTL(true);
        // and a second shape that reads the other way, to pin the arm next to it
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('AAAA');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        // The writer wrote `style:writing-mode="rl-tb"`, so it has to come back right to left
        self::assertTrue($arrayShape[0]->getParagraph()->getAlignment()->isRTL());
        self::assertInstanceOf(RichText::class, $arrayShape[1]);
        self::assertFalse($arrayShape[1]->getParagraph()->getAlignment()->isRTL());
    }

    public function testHyperlinkToUrl(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oRun = $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('AAAA');
        $oRun->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        $arrayRichText = $arrayShape[0]->getParagraph()->getRichTextElements();
        self::assertInstanceOf(TextElement::class, $arrayRichText[0]);
        $oHyperlink = $arrayRichText[0]->getHyperlink();

        self::assertFalse($oHyperlink->isInternal());
        self::assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oHyperlink->getUrl());
    }

    /**
     * @return array<array{string, null|bool}>
     */
    public static function dataProviderWritingModes(): array
    {
        return [
            // the three ODF spells right to left with, and the three it spells left to right with
            ['rl-tb', true],
            ['tb-rl', true],
            ['rl', true],
            ['lr-tb', false],
            ['tb-lr', false],
            ['lr', false],
            // and the two that state no direction of their own
            ['tb', null],
            ['page', null],
        ];
    }

    /**
     * @dataProvider dataProviderWritingModes
     */
    #[DataProvider('dataProviderWritingModes')]
    public function testTextColumnsRTLWritingModes(string $writingMode, ?bool $expected): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('AAAA');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);

        // This library writes two of the eight spellings; putting the others into a file it wrote
        // says what the reader makes of one that came from somewhere else.
        $oZip = new ZipArchive();
        $oZip->open($file);
        $sContent = $oZip->getFromName('content.xml');
        self::assertIsString($sContent);
        $oZip->deleteName('content.xml');
        $oZip->addFromString('content.xml', str_replace(
            'fo:wrap-option="wrap" style:writing-mode="lr-tb"',
            'fo:wrap-option="wrap" style:writing-mode="' . $writingMode . '"',
            $sContent
        ));
        $oZip->close();

        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertSame($expected, $arrayShape[0]->hasColumnsRTL());
    }

    public function testSlideBackgroundSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->setBackground(
            (new BackgroundColor())->setColor(new Color('FF4672A8'))
        );
        // a second slide that was given no background, so the two do not write the same style
        $oPhpPresentation->createSlide();

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $oBackground = $oPhpPresentationRead->getSlide(0)->getBackground();
        self::assertInstanceOf(BackgroundColor::class, $oBackground);
        self::assertEquals('4672A8', $oBackground->getColor()->getRGB());
        self::assertNotInstanceOf(BackgroundColor::class, $oPhpPresentationRead->getSlide(1)->getBackground());
    }

    public function testTextColumnsRTL(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('AAAA');
        $oShape->setColumns(2)->setColumnsRTL(true);
        // and a second shape that never said anything about the order of its columns
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('BBBB');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertTrue($arrayShape[0]->hasColumnsRTL());
        // the frame always states its writing mode, so what was never set comes back said outright
        self::assertInstanceOf(RichText::class, $arrayShape[1]);
        self::assertFalse($arrayShape[1]->hasColumnsRTL());
    }

    public function testTextColumns(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('AAAA');
        $oShape->setColumns(3)->setColumnSpacing(20);
        // and a second shape that never asked for columns, to pin the default
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('BBBB');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertEquals(3, $arrayShape[0]->getColumns());
        // the gap survives the trip through `fo:column-gap`, which is a length in centimetres
        self::assertEquals(20, $arrayShape[0]->getColumnSpacing());
        self::assertInstanceOf(RichText::class, $arrayShape[1]);
        self::assertEquals(1, $arrayShape[1]->getColumns());
        self::assertEquals(0, $arrayShape[1]->getColumnSpacing());
    }

    public function testTableSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oTable = $oPhpPresentation->getActiveSlide()->createTableShape(2);
        $oTable->setWidth(400)->setHeight(120)->setOffsetX(10)->setOffsetY(20);

        $oRowHeader = $oTable->createRow();
        $oRowHeader->setHeight(30);
        $oRowHeader->getCell(0)->createTextRun('Header')->getFont()->setBold(true);
        $oRowHeader->getCell(1)->createTextRun('Second');
        $oRowHeader->getCell(0)->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFEEEEEE'));

        $oRowBody = $oTable->createRow();
        $oRowBody->setHeight(40);
        $oRowBody->getCell(0)->createTextRun('Alpha');
        $oRowBody->getCell(1)->createTextRun('Beta');
        $oRowBody->getCell(1)->getBorders()->getBottom()
            ->setLineStyle(Border::LINE_SINGLE)
            ->setDashStyle(Border::DASH_DASH)
            ->setColor(new Color('FFFF0000'));

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertCount(1, $arrayShape);
        $oTableRead = $arrayShape[0];
        self::assertInstanceOf(Table::class, $oTableRead);

        self::assertEquals(2, $oTableRead->getNumColumns());
        self::assertCount(2, $oTableRead->getRows());
        self::assertEquals(400, $oTableRead->getWidth());
        self::assertEquals(10, $oTableRead->getOffsetX());
        self::assertEquals(20, $oTableRead->getOffsetY());

        self::assertEquals(30, $oTableRead->getRow(0)->getHeight());
        self::assertEquals(40, $oTableRead->getRow(1)->getHeight());

        self::assertEquals('Header', $oTableRead->getRow(0)->getCell(0)->getPlainText());
        self::assertEquals('Second', $oTableRead->getRow(0)->getCell(1)->getPlainText());
        self::assertEquals('Alpha', $oTableRead->getRow(1)->getCell(0)->getPlainText());
        self::assertEquals('Beta', $oTableRead->getRow(1)->getCell(1)->getPlainText());

        $oFont = $oTableRead->getRow(0)->getCell(0)->getParagraph()->getRichTextElements()[0]->getFont();
        self::assertInstanceOf(Font::class, $oFont);
        self::assertTrue($oFont->isBold());

        $oFill = $oTableRead->getRow(0)->getCell(0)->getFill();
        self::assertEquals(Fill::FILL_SOLID, $oFill->getFillType());
        self::assertEquals('EEEEEE', $oFill->getStartColor()->getRGB());

        $oBorder = $oTableRead->getRow(1)->getCell(1)->getBorders()->getBottom();
        self::assertEquals(Border::LINE_SINGLE, $oBorder->getLineStyle());
        self::assertEquals(Border::DASH_DASH, $oBorder->getDashStyle());
        self::assertInstanceOf(Color::class, $oBorder->getColor());
        self::assertEquals('FF0000', $oBorder->getColor()->getRGB());
    }

    /**
     * @dataProvider dataProviderFirstRow
     */
    #[DataProvider('dataProviderFirstRow')]
    public function testTableHeaderRowSurvivesTheRoundTrip(bool $firstRow): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oTable = $oPhpPresentation->getActiveSlide()->createTableShape(1);
        $oTable->setFirstRow($firstRow);
        $oTable->createRow()->getCell(0)->createTextRun('Header');
        $oTable->createRow()->getCell(0)->createTextRun('Body');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        $oTableRead = $arrayShape[0];
        self::assertInstanceOf(Table::class, $oTableRead);
        self::assertEquals($firstRow, $oTableRead->isFirstRow());

        // Every row is a row of the table, so the order the table was written in is the order it
        // reads back in
        self::assertCount(2, $oTableRead->getRows());
        self::assertEquals('Header', $oTableRead->getRow(0)->getCell(0)->getPlainText());
        self::assertEquals('Body', $oTableRead->getRow(1)->getCell(0)->getPlainText());
    }

    /**
     * @return array<array<bool>>
     */
    public static function dataProviderFirstRow(): iterable
    {
        yield [true];

        yield [false];
    }

    public function testTableWrittenByLibreOffice(): void
    {
        // A producer other than this library says a header row with `table:use-first-row-styles`
        // rather than by wrapping it, names the style of a cell on the row it belongs to, and puts
        // a replacement image of the table in the same frame -- which is not what the frame holds
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Issue_00141.odp';
        $oPhpPresentation = (new ODPresentation())->load($file);

        $arrayShape = array_values((array) $oPhpPresentation->getSlide(2)->getShapeCollection());
        self::assertCount(2, $arrayShape);
        $oTable = $arrayShape[1];
        self::assertInstanceOf(Table::class, $oTable);

        self::assertTrue($oTable->isFirstRow());
        self::assertEquals(3, $oTable->getNumColumns());
        self::assertCount(3, $oTable->getRows());
        foreach ([['1', '2', '3'], ['a', 'b', 'c'], ['A', 'B', 'C']] as $rowIndex => $expected) {
            foreach ($expected as $cellIndex => $text) {
                self::assertEquals($text, $oTable->getRow($rowIndex)->getCell($cellIndex)->getPlainText());
            }
        }
    }

    public function testFieldSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oParagraph = $oPhpPresentation->getActiveSlide()->createRichTextShape()->getActiveParagraph();
        $oParagraph->createTextRun('page ');
        $oParagraph->createField(Field::TYPE_SLIDENUM, '<nr.>');
        $oParagraph->createTextRun(' of ');
        $oParagraph->createField(Field::TYPE_SLIDECOUNT, '12');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        $arrayElements = $arrayShape[0]->getParagraph()->getRichTextElements();
        self::assertCount(4, $arrayElements);

        self::assertInstanceOf(Field::class, $arrayElements[1]);
        self::assertEquals(Field::TYPE_SLIDENUM, $arrayElements[1]->getType());
        // the text a field carries is what it stands in for, and it belongs to the field rather
        // than to the span around it
        self::assertEquals('<nr.>', $arrayElements[1]->getText());

        self::assertInstanceOf(Field::class, $arrayElements[3]);
        self::assertEquals(Field::TYPE_SLIDECOUNT, $arrayElements[3]->getType());

        self::assertNotInstanceOf(Field::class, $arrayElements[2]);
        self::assertEquals(' of ', $arrayElements[2]->getText());
    }

    /**
     * @return array<array<int|string>>
     */
    public static function dataProviderLineSpacing(): array
    {
        return [
            [Paragraph::LINE_SPACING_MODE_PERCENT, 150],
            [Paragraph::LINE_SPACING_MODE_POINT, 22],
        ];
    }

    /**
     * @dataProvider dataProviderLineSpacing
     */
    #[DataProvider('dataProviderLineSpacing')]
    public function testLineSpacingSurvivesTheRoundTrip(string $mode, int $spacing): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Sample');
        $oShape->getActiveParagraph()->setLineSpacingMode($mode)->setLineSpacing($spacing);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);

        // the line height was read out of the margin below the paragraph rather than out of
        // itself, so every spacing came back as the margin in centimetres, rounded down
        self::assertEquals($mode, $arrayShape[0]->getParagraph()->getLineSpacingMode());
        self::assertEquals($spacing, $arrayShape[0]->getParagraph()->getLineSpacing());
    }

    public function testParagraphSpacingSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Sample');
        $oShape->getActiveParagraph()->setSpacingBefore(11)->setSpacingAfter(13);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);

        // a spacing is points here and centimetres in the file, and three decimals of a
        // centimetre were not enough to give a whole number of points back
        self::assertEquals(11, $arrayShape[0]->getParagraph()->getSpacingBefore());
        self::assertEquals(13, $arrayShape[0]->getParagraph()->getSpacingAfter());
    }

    public function testShapeBorderSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Sample');
        $oShape->getBorder()
            ->setLineStyle(Border::LINE_SINGLE)
            ->setDashStyle(Border::DASH_LARGEDASHDOT)
            ->setLineWidth(4)
            ->setColor(new Color('FFFF0000'));

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);

        // the graphic style of a shape says its border as a stroke, and nothing read it back
        $oBorder = $arrayShape[0]->getBorder();
        self::assertEquals(Border::LINE_SINGLE, $oBorder->getLineStyle());
        self::assertEquals(Border::DASH_LARGEDASHDOT, $oBorder->getDashStyle());
        self::assertEquals(4, $oBorder->getLineWidth());
        self::assertEquals('FFFF0000', $oBorder->getColor()->getARGB());
    }

    public function testShapeWithoutABorderSaysSoAfterTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('Sample');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new ODPresentationWriter($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new ODPresentation())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertEquals(Border::LINE_NONE, $arrayShape[0]->getBorder()->getLineStyle());
    }
}
