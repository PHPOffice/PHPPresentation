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
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Reader\ODPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Font;
use PHPUnit\Framework\TestCase;

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
        self::assertIsArray($oPhpPresentation->getDocumentProperties()->getCustomProperties());
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
        self::assertIsArray($oPhpPresentation->getDocumentProperties()->getCustomProperties());
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
        self::assertIsArray($oPhpPresentation->getDocumentProperties()->getCustomProperties());
        self::assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        self::assertEquals('MaDiapo', $oPhpPresentation->getSlide(0)->getName());
    }

    public function testIssue00141(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Issue_00141.odp';
        $object = new ODPresentation();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        self::assertIsArray($oPhpPresentation->getDocumentProperties()->getCustomProperties());
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
}
