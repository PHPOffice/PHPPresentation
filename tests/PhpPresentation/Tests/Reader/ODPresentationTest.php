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
        $this->assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
        $this->assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $this->assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.odp';
        $this->assertTrue($object->canRead($file));
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
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        // Document Properties
        $this->assertEquals('PHPOffice', $oPhpPresentation->getDocumentProperties()->getCreator());
        $this->assertEquals('PHPPresentation Team', $oPhpPresentation->getDocumentProperties()->getLastModifiedBy());
        $this->assertEquals('Sample 02 Title', $oPhpPresentation->getDocumentProperties()->getTitle());
        $this->assertEquals('Sample 02 Subject', $oPhpPresentation->getDocumentProperties()->getSubject());
        $this->assertEquals('Sample 02 Description', $oPhpPresentation->getDocumentProperties()->getDescription());
        $this->assertEquals('office 2007 openxml libreoffice odt php', $oPhpPresentation->getDocumentProperties()->getKeywords());
        $this->assertIsArray($oPhpPresentation->getDocumentProperties()->getCustomProperties());
        $this->assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        // Presentation Properties
        $this->assertEquals(PresentationProperties::SLIDESHOW_TYPE_PRESENT, $oPhpPresentation->getPresentationProperties()->getSlideshowType());

        $this->assertCount(4, $oPhpPresentation->getAllSlides());

        // Slide 1
        $oSlide1 = $oPhpPresentation->getSlide(0);
        $arrayShape = (array) $oSlide1->getShapeCollection();
        $this->assertCount(2, $arrayShape);
        // Slide 1 : Shape 1
        /** @var Gd $oShape */
        $oShape = $arrayShape[0];
        $this->assertInstanceOf(Gd::class, $oShape);
        $this->assertEquals('PHPPresentation logo', $oShape->getName());
        $this->assertEquals('PHPPresentation logo', $oShape->getDescription());
        $this->assertEquals(36, $oShape->getHeight());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(10, $oShape->getOffsetY());
        $this->assertEquals('image/png', $oShape->getMimeType());
        $this->assertTrue($oShape->getShadow()->isVisible());
        $this->assertEquals(45, $oShape->getShadow()->getDirection());
        $this->assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 1 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        $this->assertInstanceOf(RichText::class, $oShape);
        $this->assertEquals(200, $oShape->getHeight());
        $this->assertEquals(600, $oShape->getWidth());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(400, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        $this->assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $oParagraph);
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
        $this->assertEquals(0, $oParagraph->getSpacingAfter());
        $this->assertEquals(0, $oParagraph->getSpacingBefore());
        $this->assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        $this->assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(3, $arrayRichText);
        // Slide 1 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Introduction to', $oRichText->getText());
        $this->assertTrue($oRichText->getFont()->isBold());
        $this->assertEquals(28, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 1 : Shape 2 : Paragraph 2
        $oRichText = $arrayRichText[1];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $oRichText);
        // Slide 1 : Shape 2 : Paragraph 3
        $oRichText = $arrayRichText[2];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('PHPPresentation', $oRichText->getText());
        $this->assertTrue($oRichText->getFont()->isBold());
        $this->assertEquals(60, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());

        // Slide 2
        $oSlide2 = $oPhpPresentation->getSlide(1);
        $arrayShape = (array) $oSlide2->getShapeCollection();
        $this->assertCount(3, $arrayShape);
        // Slide 2 : Shape 1
        /** @var Gd $oShape */
        $oShape = $arrayShape[0];
        $this->assertInstanceOf(Gd::class, $oShape);
        $this->assertEquals('PHPPresentation logo', $oShape->getName());
        $this->assertEquals('PHPPresentation logo', $oShape->getDescription());
        $this->assertEquals(36, $oShape->getHeight());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(10, $oShape->getOffsetY());
        $this->assertEquals('image/png', $oShape->getMimeType());
        $this->assertTrue($oShape->getShadow()->isVisible());
        $this->assertEquals(45, $oShape->getShadow()->getDirection());
        $this->assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 2 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        $this->assertInstanceOf(RichText::class, $oShape);
        $this->assertEquals(100, $oShape->getHeight());
        $this->assertEquals(930, $oShape->getWidth());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        $this->assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        // Slide 2 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('What is PHPPresentation?', $oRichText->getText());
        $this->assertTrue($oRichText->getFont()->isBold());
        $this->assertEquals(48, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 2 : Shape 3
        /** @var RichText $oShape */
        $oShape = $arrayShape[2];
        $this->assertInstanceOf(RichText::class, $oShape);
        $this->assertEquals(600, $oShape->getHeight());
        $this->assertEquals(930, $oShape->getWidth());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        $this->assertCount(4, $arrayParagraphs);
        // Slide 2 : Shape 3 : Paragraph 1
        $oParagraph = $arrayParagraphs[0];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('A class library', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 2 : Shape 3 : Paragraph 2
        $oParagraph = $arrayParagraphs[1];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Written in PHP', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 2 : Shape 3 : Paragraph 3
        $oParagraph = $arrayParagraphs[2];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Representing a presentation', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 2 : Shape 3 : Paragraph 4
        $oParagraph = $arrayParagraphs[3];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Supports writing to different file formats', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());

        // Slide 3
        $oSlide2 = $oPhpPresentation->getSlide(2);
        $arrayShape = (array) $oSlide2->getShapeCollection();
        $this->assertCount(3, $arrayShape);
        // Slide 3 : Shape 1
        /** @var Gd $oShape */
        $oShape = $arrayShape[0];
        $this->assertInstanceOf(Gd::class, $oShape);
        $this->assertEquals('PHPPresentation logo', $oShape->getName());
        $this->assertEquals('PHPPresentation logo', $oShape->getDescription());
        $this->assertEquals(36, $oShape->getHeight());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(10, $oShape->getOffsetY());
        $this->assertEquals('image/png', $oShape->getMimeType());
        $this->assertTrue($oShape->getShadow()->isVisible());
        $this->assertEquals(45, $oShape->getShadow()->getDirection());
        $this->assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 3 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        $this->assertInstanceOf(RichText::class, $oShape);
        $this->assertEquals(100, $oShape->getHeight());
        $this->assertEquals(930, $oShape->getWidth());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        $this->assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
        $this->assertEquals(0, $oParagraph->getSpacingAfter());
        $this->assertEquals(0, $oParagraph->getSpacingBefore());
        $this->assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        $this->assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        // Slide 3 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('What\'s the point?', $oRichText->getText());
        $this->assertTrue($oRichText->getFont()->isBold());
        $this->assertEquals(48, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 3 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[2];
        $this->assertInstanceOf(RichText::class, $oShape);
        $this->assertEquals(600, $oShape->getHeight());
        $this->assertEquals(930, $oShape->getWidth());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        $this->assertCount(8, $arrayParagraphs);
        // Slide 3 : Shape 3 : Paragraph 1
        $oParagraph = $arrayParagraphs[0];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(0, $oParagraph->getAlignment()->getLevel());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Generate slide decks', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 3 : Shape 3 : Paragraph 2
        $oParagraph = $arrayParagraphs[1];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(1, $oParagraph->getAlignment()->getLevel());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Represent business data', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 3 : Shape 3 : Paragraph 3
        $oParagraph = $arrayParagraphs[2];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(1, $oParagraph->getAlignment()->getLevel());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Show a family slide show', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 3 : Shape 3 : Paragraph 4
        $oParagraph = $arrayParagraphs[3];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(1, $oParagraph->getAlignment()->getLevel());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('...', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 3 : Shape 3 : Paragraph 5
        $oParagraph = $arrayParagraphs[4];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(0, $oParagraph->getAlignment()->getLevel());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Export these to different formats', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 3 : Shape 3 : Paragraph 6
        $oParagraph = $arrayParagraphs[5];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(1, $oParagraph->getAlignment()->getLevel());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('PHPPresentation 2007', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 3 : Shape 3 : Paragraph 7
        $oParagraph = $arrayParagraphs[6];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(1, $oParagraph->getAlignment()->getLevel());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Serialized', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 3 : Shape 3 : Paragraph 8
        $oParagraph = $arrayParagraphs[7];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
//         $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
//         $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        $this->assertEquals(1, $oParagraph->getAlignment()->getLevel());
        $this->assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('... (more to come) ...', $oRichText->getText());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());

        // Slide 4
        $oSlide3 = $oPhpPresentation->getSlide(3);
        $arrayShape = (array) $oSlide3->getShapeCollection();
        $this->assertCount(3, $arrayShape);
        // Slide 4 : Shape 1
        /** @var Gd $oShape */
        $oShape = $arrayShape[0];
        $this->assertInstanceOf(Gd::class, $oShape);
        $this->assertEquals('PHPPresentation logo', $oShape->getName());
        $this->assertEquals('PHPPresentation logo', $oShape->getDescription());
        $this->assertEquals(36, $oShape->getHeight());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(10, $oShape->getOffsetY());
        $this->assertEquals('image/png', $oShape->getMimeType());
        $this->assertTrue($oShape->getShadow()->isVisible());
        $this->assertEquals(45, $oShape->getShadow()->getDirection());
        $this->assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 4 : Shape 2
        /** @var RichText $oShape */
        $oShape = $arrayShape[1];
        $this->assertInstanceOf(RichText::class, $oShape);
        $this->assertEquals(100, $oShape->getHeight());
        $this->assertEquals(930, $oShape->getWidth());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        $this->assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
        $this->assertEquals(0, $oParagraph->getSpacingAfter());
        $this->assertEquals(0, $oParagraph->getSpacingBefore());
        $this->assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        $this->assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayRichText);
        // Slide 4 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Need more info?', $oRichText->getText());
        $this->assertTrue($oRichText->getFont()->isBold());
        $this->assertEquals(48, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 4 : Shape 3
        /** @var RichText $oShape */
        $oShape = $arrayShape[2];
        $this->assertInstanceOf(RichText::class, $oShape);
        $this->assertEquals(600, $oShape->getHeight());
        $this->assertEquals(930, $oShape->getWidth());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        $this->assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $this->assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $this->assertFalse($oParagraph->getAlignment()->isRTL());
        $this->assertEquals(0, $oParagraph->getSpacingAfter());
        $this->assertEquals(0, $oParagraph->getSpacingBefore());
        $this->assertEquals(Paragraph::LINE_SPACING_MODE_PERCENT, $oParagraph->getLineSpacingMode());
        $this->assertEquals(100, $oParagraph->getLineSpacing());
        $arrayRichText = $oParagraph->getRichTextElements();
        $this->assertCount(3, $arrayRichText);
        // Slide 4 : Shape 3 : Paragraph 1
        $oRichText = $arrayRichText[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        $this->assertEquals('Check the project site on GitHub:', $oRichText->getText());
        $this->assertFalse($oRichText->getFont()->isBold());
        $this->assertEquals(36, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        // Slide 4 : Shape 3 : Paragraph 2
        $oRichText = $arrayRichText[1];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $oRichText);
        // Slide 4 : Shape 3 : Paragraph 3
        /** @var RichText\Run $oRichText */
        $oRichText = $arrayRichText[2];
        $this->assertInstanceOf(RichText\Run::class, $oRichText);
        $this->assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getText());
        $this->assertFalse($oRichText->getFont()->isBold());
        $this->assertEquals(32, $oRichText->getFont()->getSize());
        $this->assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        $this->assertEquals('Calibri', $oRichText->getFont()->getName());
        $this->assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        $this->assertTrue($oRichText->hasHyperlink());
        $this->assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getHyperlink()->getUrl());
        //$this->assertEquals('PHPPresentation', $oRichText->getHyperlink()->getTooltip());
    }

    public function testSlideName(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/ODP_Slide_Name.odp';
        $object = new ODPresentation();
        $oPhpPresentation = $object->load($file);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        $this->assertIsArray($oPhpPresentation->getDocumentProperties()->getCustomProperties());
        $this->assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        $this->assertEquals('MaDiapo', $oPhpPresentation->getSlide(0)->getName());
    }

    public function testIssue00141(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Issue_00141.odp';
        $object = new ODPresentation();
        $oPhpPresentation = $object->load($file);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        $this->assertIsArray($oPhpPresentation->getDocumentProperties()->getCustomProperties());
        $this->assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        $this->assertCount(3, $oPhpPresentation->getAllSlides());

        // Slide 1
        $oSlide = $oPhpPresentation->getSlide(1);
        $arrayShape = (array) $oSlide->getShapeCollection();
        $this->assertCount(2, $arrayShape);
        // Slide 1 : Shape 1
        /** @var RichText $oShape */
        $oShape = reset($arrayShape);
        $this->assertInstanceOf(RichText::class, $oShape);
        // Slide 1 : Shape 1 : Paragraph 1
        $oParagraph = $oShape->getParagraph();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Paragraph', $oParagraph);
        // Slide 1 : Shape 1 : Paragraph 1 : RichText Elements
        $arrayElements = $oParagraph->getRichTextElements();
        $this->assertCount(1, $arrayElements);
        $oElement = reset($arrayElements);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $oElement);
        $this->assertEquals('TEST IMAGE', $oElement->getText());
    }
}
