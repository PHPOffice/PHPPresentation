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

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Reader\PowerPoint2007;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Doughnut;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Radar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Field;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007 as PowerPoint2007Writer;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use ZipArchive;

/**
 * Test class for PowerPoint2007 reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\PowerPoint2007
 */
class PowerPoint2007Test extends TestCase
{
    /**
     * Test can read.
     */
    public function testCanRead(): void
    {
        $object = new PowerPoint2007();

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt';
        self::assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
        self::assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        self::assertTrue($object->canRead($file));
    }

    public function testLoadFileNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new PowerPoint2007();
        $object->load('');
    }

    public function testLoadFileBadFormat(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt';
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage(sprintf(
            'The file %s is not in the format supported by class PhpOffice\PhpPresentation\Reader\PowerPoint2007',
            $file
        ));

        $object = new PowerPoint2007();
        $object->load($file);
    }

    public function testFileSupportsNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new PowerPoint2007();
        $object->fileSupportsUnserializePhpPresentation('');
    }

    public function testLoadFile01(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        // Document Properties
        self::assertEquals('PHPOffice', $oPhpPresentation->getDocumentProperties()->getCreator());
        self::assertEquals('PHPPresentation Team', $oPhpPresentation->getDocumentProperties()->getLastModifiedBy());
        self::assertEquals('Sample 02 Title', $oPhpPresentation->getDocumentProperties()->getTitle());
        self::assertEquals('Sample 02 Subject', $oPhpPresentation->getDocumentProperties()->getSubject());
        self::assertEquals('Sample 02 Description', $oPhpPresentation->getDocumentProperties()->getDescription());
        self::assertEquals('office 2007 openxml libreoffice odt php', $oPhpPresentation->getDocumentProperties()->getKeywords());
        self::assertEquals('Sample Category', $oPhpPresentation->getDocumentProperties()->getCategory());
        self::assertEquals('', $oPhpPresentation->getDocumentProperties()->getRevision());
        self::assertEquals('', $oPhpPresentation->getDocumentProperties()->getStatus());
        self::assertCount(0, $oPhpPresentation->getDocumentProperties()->getCustomProperties());

        // Presentation Properties
        self::assertEquals(PresentationProperties::SLIDESHOW_TYPE_PRESENT, $oPhpPresentation->getPresentationProperties()->getSlideshowType());
        // Document Layout
        self::assertEquals(DocumentLayout::LAYOUT_SCREEN_4X3, $oPhpPresentation->getLayout()->getDocumentLayout());
        self::assertEquals(254, $oPhpPresentation->getLayout()->getCX(DocumentLayout::UNIT_MILLIMETER));
        self::assertEquals(190.5, $oPhpPresentation->getLayout()->getCY(DocumentLayout::UNIT_MILLIMETER));

        // Slides
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
        self::assertEquals('image/gif', $oShape->getMimeType());
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
        self::assertEquals('image/gif', $oShape->getMimeType());
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
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
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
        self::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals('image/gif', $oShape->getMimeType());
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
        self::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(0, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(0, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        self::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        self::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        self::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        self::assertFalse($oParagraph->getAlignment()->isRTL());
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
        self::assertEquals('image/gif', $oShape->getMimeType());
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
        self::assertTrue($oRichText->hasHyperlink());
        self::assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getHyperlink()->getUrl());
        self::assertEquals('PHPPresentation', $oRichText->getHyperlink()->getTooltip());
        self::assertFalse($oRichText->getHyperlink()->isTextColorUsed());
        self::assertEquals('Calibri', $oRichText->getFont()->getName());
        self::assertEquals(Font::FORMAT_LATIN, $oRichText->getFont()->getFormat());
        self::assertEquals(Font::CAPITALIZATION_NONE, $oRichText->getFont()->getCapitalization());
    }

    /**
     * The shapes of a group are written in the coordinate space a:chOff/a:chExt declares,
     * which is not necessarily the slide's. PPTX_Group.pptx declares one that is offset and
     * scaled, so reading it back at face value would put every shape in the wrong place.
     */
    public function testLoadFileGroup(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_Group.pptx';
        $oPhpPresentation = (new PowerPoint2007())->load($file);

        $oSlide = $oPhpPresentation->getSlide(0);
        self::assertCount(2, $oSlide->getShapeCollection());

        // The shape outside the group is untouched by any of this.
        $oFree = $oSlide->getShapeCollection()[0];
        self::assertInstanceOf(RichText::class, $oFree);
        self::assertEquals(50, $oFree->getOffsetX());
        self::assertEquals(50, $oFree->getOffsetY());

        // a:off is 200px to the right of a:chOff, and a:ext is twice a:chExt.
        $oGroup = $oSlide->getShapeCollection()[1];
        self::assertInstanceOf(Group::class, $oGroup);
        self::assertEquals(30, $oGroup->getRotation());
        self::assertEquals(500, $oGroup->getOffsetX());
        self::assertEquals(100, $oGroup->getOffsetY());
        self::assertEquals(760, $oGroup->getExtentX());
        self::assertEquals(680, $oGroup->getExtentY());
        self::assertCount(3, $oGroup->getShapeCollection());

        // Written at 300,100 and 500,260, sized 180x60.
        foreach ([[500, 100], [900, 420]] as $index => [$offsetX, $offsetY]) {
            $oShape = $oGroup->getShapeCollection()[$index];
            self::assertInstanceOf(RichText::class, $oShape);
            self::assertEquals($offsetX, $oShape->getOffsetX());
            self::assertEquals($offsetY, $oShape->getOffsetY());
            self::assertEquals(360, $oShape->getWidth());
            self::assertEquals(120, $oShape->getHeight());
        }

        // A group inside the group: the mapping reaches its shapes too.
        $oInner = $oGroup->getShapeCollection()[2];
        self::assertInstanceOf(Group::class, $oInner);
        self::assertCount(1, $oInner->getShapeCollection());
        $oShape = $oInner->getShapeCollection()[0];
        self::assertInstanceOf(RichText::class, $oShape);
        self::assertEquals(700, $oShape->getOffsetX());
        self::assertEquals(700, $oShape->getOffsetY());
        self::assertEquals(200, $oShape->getWidth());
        self::assertEquals(80, $oShape->getHeight());
    }

    /**
     * PPTX_GroupNested.pptx comes out of PowerPoint 16 and holds a group inside a group.
     * The outer one declares the coordinate space of the slide; the inner one was resized
     * after it was made, so it declares one that is scaled by about eight in x and two in
     * y, and the shapes inside it are still written at the size they were drawn.
     */
    public function testLoadFileGroupNested(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_GroupNested.pptx';
        $oPhpPresentation = (new PowerPoint2007())->load($file);

        $oSlide = $oPhpPresentation->getSlide(0);
        self::assertCount(1, $oSlide->getShapeCollection());

        $oOuter = $oSlide->getShapeCollection()[0];
        self::assertInstanceOf(Group::class, $oOuter);
        self::assertCount(2, $oOuter->getShapeCollection());

        // Written at 252,55 in a group that declares the space it is already in.
        $oTriangle = $oOuter->getShapeCollection()[1];
        self::assertInstanceOf(RichText::class, $oTriangle);
        self::assertEquals(252, $oTriangle->getOffsetX());
        self::assertEquals(54, $oTriangle->getOffsetY());

        // Written at 523,199 and 549,334, sized 96x61 and 81x69.
        $oInner = $oOuter->getShapeCollection()[0];
        self::assertInstanceOf(Group::class, $oInner);
        self::assertCount(2, $oInner->getShapeCollection());
        foreach ([[240, 198, 762, 111], [446, 443, 643, 125]] as $index => [$offsetX, $offsetY, $width, $height]) {
            $oShape = $oInner->getShapeCollection()[$index];
            self::assertInstanceOf(RichText::class, $oShape);
            self::assertEquals($offsetX, $oShape->getOffsetX());
            self::assertEquals($offsetY, $oShape->getOffsetY());
            self::assertEquals($width, $oShape->getWidth());
            self::assertEquals($height, $oShape->getHeight());
        }
    }

    public function testLoadFileChartBar(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_ChartBar.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        // Document Properties

        // Slides
        self::assertCount(3, $oPhpPresentation->getAllSlides());

        // Slide 2
        $oSlide2 = $oPhpPresentation->getSlide(1);
        $arrayShape = (array) $oSlide2->getShapeCollection();
        self::assertCount(2, $arrayShape);
        // Slide 2 : Shape 2
        /** @var Chart $oShape */
        $oShape = $arrayShape[1];
        self::assertInstanceOf(Chart::class, $oShape);
        self::assertInstanceOf(Bar::class, $oShape->getPlotArea()->getType());
    }

    public function testLoadFileWithoutImages(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file, PowerPoint2007::SKIP_IMAGES);
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
        // $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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
        // $this->assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        // $this->assertEquals(-25, $oParagraph->getAlignment()->getIndent());
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

    public function testMarkAsFinal(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        self::assertFalse($oPhpPresentation->getPresentationProperties()->isMarkedAsFinal());

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_MarkAsFinal.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        self::assertTrue($oPhpPresentation->getPresentationProperties()->isMarkedAsFinal());
    }

    public function testZoom(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        self::assertEquals(1, $oPhpPresentation->getPresentationProperties()->getZoom());

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_Zoom.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        self::assertEquals(2.68, $oPhpPresentation->getPresentationProperties()->getZoom());
    }

    public function testSlideLayout(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Issue_00150.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);

        $masterSlides = $oPhpPresentation->getAllMasterSlides();
        self::assertCount(3, $masterSlides);
        self::assertCount(11, $masterSlides[0]->getAllSlideLayouts());
        self::assertCount(11, $masterSlides[1]->getAllSlideLayouts());
        self::assertCount(11, $masterSlides[2]->getAllSlideLayouts());
    }

    public function testLoadFileWithInvalidImages(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_InvalidImage.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        self::assertEquals(1, $oPhpPresentation->getSlideCount());
    }

    public function testSeriesDataPoints(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSeries = new Series('Downloads', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        $oSeries->getDataPointOutline(0)->setWidth(2)->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_WHITE));
        $oSeries->getDataPointFill(1)->setFillType(Fill::FILL_NONE);
        $oSeries->getDataPointOutline(1)->getFill()->setFillType(Fill::FILL_NONE);
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oPhpPresentation->getActiveSlide()->createChartShape()->getPlotArea()->setType($oBar);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $seriesRead = $arrayShape[0]->getPlotArea()->getType()->getSeries();
        $seriesRead = reset($seriesRead);

        self::assertEquals(Fill::FILL_SOLID, $seriesRead->getDataPointFill(0)->getFillType());
        self::assertEquals('0000FF', $seriesRead->getDataPointFill(0)->getStartColor()->getRGB());
        self::assertEquals(Fill::FILL_SOLID, $seriesRead->getDataPointOutline(0)->getFill()->getFillType());
        self::assertEquals('FFFFFF', $seriesRead->getDataPointOutline(0)->getFill()->getStartColor()->getRGB());
        self::assertEquals(2, $seriesRead->getDataPointOutline(0)->getWidth());

        self::assertEquals(Fill::FILL_NONE, $seriesRead->getDataPointFill(1)->getFillType());
        self::assertEquals(Fill::FILL_NONE, $seriesRead->getDataPointOutline(1)->getFill()->getFillType());
    }

    /**
     * @dataProvider dataProviderColorAlpha
     */
    #[DataProvider('dataProviderColorAlpha')]
    public function testColorAlphaIsReadBack(string $argb, int $expectedAlpha): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSeries = new Series('Downloads', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($argb));
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oPhpPresentation->getActiveSlide()->createChartShape()->getPlotArea()->setType($oBar);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $seriesRead = $arrayShape[0]->getPlotArea()->getType()->getSeries();
        $seriesRead = reset($seriesRead);

        // the colour comes back whole: an alpha that is not `FF` used to be read as `0`, and as a
        // single character it took the first digit of the colour with it
        self::assertEquals($argb, $seriesRead->getFill()->getStartColor()->getARGB());
        self::assertEquals(substr($argb, 2), $seriesRead->getFill()->getStartColor()->getRGB());
        self::assertEquals($expectedAlpha, $seriesRead->getFill()->getStartColor()->getAlpha());
    }

    /**
     * @return array<array{string, int}>
     */
    public static function dataProviderColorAlpha(): array
    {
        return [
            // ARGB written, and the percent it is read back as
            ['FFFF7700', 100],
            // 84%: the plate a data label sits on
            ['D6FFFFFF', 84],
            // an alpha of one hex digit, which has to keep its leading zero
            ['0A00FF00', 4],
        ];
    }

    public function testSeriesLabelFillIsReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSeries = new Series('Downloads', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        // the plate the labels sit on: white, and see-through enough to read the chart under it
        $oSeries->getLabelFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('D6FFFFFF'));
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oPhpPresentation->getActiveSlide()->createChartShape()->getPlotArea()->setType($oBar);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $seriesRead = $arrayShape[0]->getPlotArea()->getType()->getSeries();
        $seriesRead = reset($seriesRead);

        self::assertEquals(Fill::FILL_SOLID, $seriesRead->getLabelFill()->getFillType());
        self::assertEquals('FFFFFF', $seriesRead->getLabelFill()->getStartColor()->getRGB());
        self::assertEquals(84, $seriesRead->getLabelFill()->getStartColor()->getAlpha());
        // the serie keeps its own fill, which is not the one behind the labels
        self::assertNotEquals(Fill::FILL_SOLID, $seriesRead->getFill()->getFillType());
    }

    public function testSeriesDataPointPatternFillIsReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSeries = new Series('Downloads', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']);
        $oSeries->getDataPointFill(0)
            ->setFillType(Fill::FILL_PATTERN_WDDNDIAG)
            ->setStartColor(new Color(Color::COLOR_BLUE))
            ->setEndColor(new Color(Color::COLOR_YELLOW));
        // A pattern the writer leaves without `prst`, because it is not one: the element that comes
        // out carries two colours and names no pattern, and there is no fill type to read it back
        // as. It has to arrive as no fill at all rather than as the string it went in as.
        $oSeries->getDataPointFill(1)->setFillType('notAPattern');
        $oBar = new Bar();
        $oBar->addSeries($oSeries);
        $oPhpPresentation->getActiveSlide()->createChartShape()->getPlotArea()->setType($oBar);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $oBarRead = $arrayShape[0]->getPlotArea()->getType();
        self::assertInstanceOf(Bar::class, $oBarRead);
        $seriesRead = $oBarRead->getSeries();
        $seriesRead = reset($seriesRead);

        self::assertEquals(Fill::FILL_PATTERN_WDDNDIAG, $seriesRead->getDataPointFill(0)->getFillType());
        self::assertEquals('0000FF', $seriesRead->getDataPointFill(0)->getStartColor()->getRGB());
        self::assertEquals('FFFF00', $seriesRead->getDataPointFill(0)->getEndColor()->getRGB());

        self::assertEquals(Fill::FILL_NONE, $seriesRead->getDataPointFill(1)->getFillType());
    }

    /**
     * @param class-string<AbstractType> $className
     *
     * @dataProvider dataProviderChartTypes
     */
    #[DataProvider('dataProviderChartTypes')]
    public function testChartTypeIsReadBack(string $className): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oType = new $className();
        $oType->addSeries(new Series('Downloads', ['Jan' => '1', 'Feb' => '5', 'Mar' => '2']));
        $oPhpPresentation->getActiveSlide()->createChartShape()->getPlotArea()->setType($oType);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $oTypeRead = $arrayShape[0]->getPlotArea()->getType();
        self::assertInstanceOf($className, $oTypeRead);

        $seriesRead = $oTypeRead->getSeries();
        self::assertCount(1, $seriesRead);
        $seriesRead = reset($seriesRead);
        self::assertEquals('Downloads', $seriesRead->getTitle());
        self::assertEquals(['Jan' => '1', 'Feb' => '5', 'Mar' => '2'], $seriesRead->getValues());
    }

    /**
     * @return array<int, array<int, string>>
     */
    public static function dataProviderChartTypes(): array
    {
        return [
            [Area::class],
            [Bar::class],
            [Bar3D::class],
            [Doughnut::class],
            [Line::class],
            [Pie::class],
            [Pie3D::class],
            [Radar::class],
            [Scatter::class],
        ];
    }

    public function testBarSettingsAreReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oBar = new Bar();
        $oBar->setBarDirection(Bar::DIRECTION_HORIZONTAL)
            ->setBarGrouping(Bar::GROUPING_STACKED)
            ->setGapWidthPercent(42)
            ->setOverlapWidthPercent(17);
        $oBar->addSeries(new Series('Downloads', ['Jan' => '1', 'Feb' => '5']));
        $oPhpPresentation->getActiveSlide()->createChartShape()->getPlotArea()->setType($oBar);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $oBarRead = $arrayShape[0]->getPlotArea()->getType();
        self::assertInstanceOf(Bar::class, $oBarRead);
        self::assertEquals(Bar::DIRECTION_HORIZONTAL, $oBarRead->getBarDirection());
        self::assertEquals(Bar::GROUPING_STACKED, $oBarRead->getBarGrouping());
        self::assertEquals(42, $oBarRead->getGapWidthPercent());
        self::assertEquals(17, $oBarRead->getOverlapWidthPercent());
    }

    public function testPieSettingsAreReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oDoughnut = new Doughnut();
        $oDoughnut->setHoleSize(35)->setFirstSliceAngle(90);
        $oDoughnut->addSeries(new Series('Downloads', ['Jan' => '1', 'Feb' => '5']));
        $oPhpPresentation->getActiveSlide()->createChartShape()->getPlotArea()->setType($oDoughnut);

        // The explosion of a slice is written by the pie in three dimensions alone, so it is the one
        // that can be read back.
        $oPie3D = new Pie3D();
        $oPie3D->setExplosion(25);
        $oPie3D->addSeries(new Series('Downloads', ['Jan' => '1', 'Feb' => '5']));
        $oPhpPresentation->createSlide()->createChartShape()->getPlotArea()->setType($oPie3D);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $oDoughnutRead = $arrayShape[0]->getPlotArea()->getType();
        self::assertInstanceOf(Doughnut::class, $oDoughnutRead);
        self::assertEquals(35, $oDoughnutRead->getHoleSize());
        self::assertEquals(90, $oDoughnutRead->getFirstSliceAngle());

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(1)->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $oPie3DRead = $arrayShape[0]->getPlotArea()->getType();
        self::assertInstanceOf(Pie3D::class, $oPie3DRead);
        self::assertEquals(25, $oPie3DRead->getExplosion());
    }

    public function testLineSmoothIsReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oLine = new Line();
        $oLine->setIsSmooth(true);
        $oLine->addSeries(new Series('Downloads', ['Jan' => '1', 'Feb' => '5']));
        $oPhpPresentation->getActiveSlide()->createChartShape()->getPlotArea()->setType($oLine);

        $oLine = new Line();
        $oLine->setIsSmooth(false);
        $oLine->addSeries(new Series('Downloads', ['Jan' => '1', 'Feb' => '5']));
        $oPhpPresentation->createSlide()->createChartShape()->getPlotArea()->setType($oLine);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(0)->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $oLineRead = $arrayShape[0]->getPlotArea()->getType();
        self::assertInstanceOf(Line::class, $oLineRead);
        self::assertTrue($oLineRead->isSmooth());

        $arrayShape = array_values((array) $oPhpPresentationRead->getSlide(1)->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $oLineRead = $arrayShape[0]->getPlotArea()->getType();
        self::assertInstanceOf(Line::class, $oLineRead);
        self::assertFalse($oLineRead->isSmooth());
    }

    public function testAxisTextIsReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oBar = new Bar();
        $oBar->addSeries(new Series('Downloads', ['Jan' => '1', 'Feb' => '5']));
        $oShape = $oPhpPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oBar);

        $oAxisX = $oShape->getPlotArea()->getAxisX();
        $oAxisX->setTitle('Quarter')->setTitleRotation(45);
        $oAxisX->getFont()->setName('Georgia')->setSize(18)->setBold(true)->setItalic(true)
            ->setUnderline(Font::UNDERLINE_DOUBLE)->getColor()->setRGB('FF0000');
        $oAxisX->getTickLabelFont()->setName('Verdana')->setSize(7)->getColor()->setRGB('00AA00');
        $oShape->getPlotArea()->getAxisY()->setTitle('Downloads');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);
        $oAxisXRead = $arrayShape[0]->getPlotArea()->getAxisX();

        self::assertEquals('Quarter', $oAxisXRead->getTitle());
        self::assertEquals(45, $oAxisXRead->getTitleRotation());

        $oFont = $oAxisXRead->getFont();
        self::assertInstanceOf(Font::class, $oFont);
        self::assertEquals('Georgia', $oFont->getName());
        self::assertEquals(18, $oFont->getSize());
        self::assertTrue($oFont->isBold());
        self::assertTrue($oFont->isItalic());
        self::assertEquals(Font::UNDERLINE_DOUBLE, $oFont->getUnderline());
        self::assertEquals('FF0000', $oFont->getColor()->getRGB());

        // the two fonts of an axis are separate: the title is styled by one, the labels by the other
        $oTickLabelFont = $oAxisXRead->getTickLabelFont();
        self::assertInstanceOf(Font::class, $oTickLabelFont);
        self::assertEquals('Verdana', $oTickLabelFont->getName());
        self::assertEquals(7, $oTickLabelFont->getSize());
        self::assertFalse($oTickLabelFont->isBold());
        self::assertEquals('00AA00', $oTickLabelFont->getColor()->getRGB());

        self::assertEquals('Downloads', $arrayShape[0]->getPlotArea()->getAxisY()->getTitle());
    }

    public function testAxisTickMarksAndOutlineAreReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oBar = new Bar();
        $oBar->addSeries(new Series('Downloads', ['Jan' => '1']));
        $oShape = $oPhpPresentation->getActiveSlide()->createChartShape();
        $oShape->getPlotArea()->setType($oBar);

        foreach ([$oShape->getPlotArea()->getAxisX(), $oShape->getPlotArea()->getAxisY()] as $oAxis) {
            $oAxis->setMajorTickMark(Axis::TICK_MARK_OUTSIDE);
            $oAxis->setMinorTickMark(Axis::TICK_MARK_INSIDE);
            $oAxis->getOutline()->setWidth(3)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->setStartColor(new Color('FF00FF00'));
        }

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(Chart::class, $arrayShape[0]);

        // both axes, because the two blocks that read them are copies of one another
        foreach ([$arrayShape[0]->getPlotArea()->getAxisX(), $arrayShape[0]->getPlotArea()->getAxisY()] as $oAxis) {
            self::assertEquals(Axis::TICK_MARK_OUTSIDE, $oAxis->getMajorTickMark());
            self::assertEquals(Axis::TICK_MARK_INSIDE, $oAxis->getMinorTickMark());
            self::assertEquals(3, $oAxis->getOutline()->getWidth());
            self::assertEquals('00FF00', $oAxis->getOutline()->getFill()->getStartColor()->getRGB());
        }
    }

    public function testShapeAskingForNoFillIsReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Nothing behind me');
        $oShape->getFill()->setFillType(Fill::FILL_NONE);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        // The shape wrote `a:noFill`, so it must come back with a fill that says so, not an unset one
        self::assertEquals(Fill::FILL_NONE, $arrayShape[0]->getFill()->getFillType());
    }

    public function testDocumentLayoutWithExplicitCustomType(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('Part 1:');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);

        // ST_SlideSizeType allows type="custom"; PowerPoint omits the attribute but other producers write it.
        // The writer emits type="screen4x3", so swapping it for "custom" says what such a file is without a binary fixture.
        $oZip = new ZipArchive();
        $oZip->open($file);
        $sPresentation = $oZip->getFromName('ppt/presentation.xml');
        self::assertIsString($sPresentation);
        $oZip->deleteName('ppt/presentation.xml');
        $oZip->addFromString(
            'ppt/presentation.xml',
            (string) preg_replace('#(<p:sldSz[^>]*type=")[^"]*(")#', '${1}custom${2}', $sPresentation)
        );
        $oZip->close();

        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        self::assertEquals(DocumentLayout::LAYOUT_CUSTOM, $oPhpPresentationRead->getLayout()->getDocumentLayout());
        self::assertEquals(9144000, $oPhpPresentationRead->getLayout()->getCX());
        self::assertEquals(6858000, $oPhpPresentationRead->getLayout()->getCY());
    }

    public function testColumnsRTL(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('AAAA');
        $oShape->setColumns(2)->setColumnsRTL(true);
        // and a second shape that never said anything about the order of its columns
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('BBBB');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertTrue($arrayShape[0]->hasColumnsRTL());
        // `rtlCol` is written for every text body, so what was never set comes back said outright
        self::assertInstanceOf(RichText::class, $arrayShape[1]);
        self::assertFalse($arrayShape[1]->hasColumnsRTL());
    }

    public function testTextRunWithoutProperties(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('Part 1:');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);

        // `a:rPr` is optional, and Keynote writes a run without one. Taking it back out of a file
        // this library wrote says what such a run is without carrying a binary fixture for it.
        $oZip = new ZipArchive();
        $oZip->open($file);
        $sSlide = $oZip->getFromName('ppt/slides/slide1.xml');
        self::assertIsString($sSlide);
        $oZip->deleteName('ppt/slides/slide1.xml');
        $oZip->addFromString(
            'ppt/slides/slide1.xml',
            (string) preg_replace('#<a:rPr[^>]*/>|<a:rPr.*?</a:rPr>#s', '', $sSlide)
        );
        $oZip->close();

        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertCount(1, $arrayShape[0]->getParagraph(0)->getRichTextElements());
        self::assertEquals('Part 1:', $arrayShape[0]->getPlainText());
    }

    public function testTextRunWithoutText(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('Part 1:');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);

        // A run without `a:t` is malformed, and the reader is not the one that should say so
        $oZip = new ZipArchive();
        $oZip->open($file);
        $sSlide = $oZip->getFromName('ppt/slides/slide1.xml');
        self::assertIsString($sSlide);
        $oZip->deleteName('ppt/slides/slide1.xml');
        $oZip->addFromString(
            'ppt/slides/slide1.xml',
            (string) preg_replace('#<a:t>.*?</a:t>#s', '', $sSlide)
        );
        $oZip->close();

        // Reading it must not raise anything of its own; the suite fails on a warning of any kind
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertEquals('', $arrayShape[0]->getPlainText());
    }

    public function testLoadingFileWithNoteInSlide(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_SlideNoteWithRichText.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
    }

    public function testShapeDecorative(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oSlide->createRichTextShape()->setDecorative()->createTextRun('Decorative');
        $oSlide->createRichTextShape()->createTextRun('Meaningful');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
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
        $oRichText->setName('Budget');
        $oRichText->setDescription('Budget spent to date: 45% of 1.2M EUR');
        $oRichText->createTextRun('45%');
        $oSlide->createRichTextShape()->createTextRun('AAA');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertCount(2, $arrayShape);
        self::assertEquals('Budget', $arrayShape[0]->getName());
        self::assertEquals('Budget spent to date: 45% of 1.2M EUR', $arrayShape[0]->getDescription());
        self::assertEquals('', $arrayShape[1]->getDescription());
    }

    /**
     * @dataProvider dataProviderCharsets
     */
    #[DataProvider('dataProviderCharsets')]
    public function testFontCharsetSurvivesTheRoundTrip(int $charset): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oRun = $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('MyText');
        $oRun->getFont()->setCharset($charset);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        $arrayElement = $arrayShape[0]->getParagraph()->getRichTextElements();

        self::assertEquals($charset, $arrayElement[0]->getFont()->getCharset());
    }

    /**
     * @return array<int, array<int, int>>
     */
    public static function dataProviderCharsets(): array
    {
        return [
            [Font::CHARSET_DEFAULT],
            [0],
            [18],
            // the ones that do not fit in a byte, which the file spells as negative
            [134],
            [178],
            [204],
            [255],
        ];
    }

    public function testSlideName(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->setName('Introduction');
        $oPhpPresentation->createSlide();

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        self::assertEquals('Introduction', $oPhpPresentationRead->getSlide(0)->getName());
        self::assertNull($oPhpPresentationRead->getSlide(1)->getName());
    }

    public function testHyperlinkToSlide(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oRun = $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('AAAA');
        $oRun->getHyperlink()->setSlideNumber(2);
        $oPhpPresentation->createSlide();

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = $oPhpPresentationRead->getSlide(0)->getShapeCollection();
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        $arrayRichText = $arrayShape[0]->getParagraph()->getRichTextElements();
        self::assertInstanceOf(TextElement::class, $arrayRichText[0]);
        $oHyperlink = $arrayRichText[0]->getHyperlink();

        // Without the slide number this comes back as the raw part name, `slide2.xml`
        self::assertTrue($oHyperlink->isInternal());
        self::assertEquals(2, $oHyperlink->getSlideNumber());
    }

    public function testFillUnsetSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Nothing was said about my fill');
        $oShape->setFill(null);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        // The shape wrote no fill of any kind, so `p:spPr` names none and the shape comes back
        // with a fill that says as much -- never a null, which no writer can read
        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertEquals(Fill::FILL_UNSET, $arrayShape[0]->getFill()->getFillType());
    }

    public function testTableFirstRowAndBandRow(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oSlide->createTableShape(1)->setFirstRow(false)->setBandRow(false)->createRow();
        $oSlide->createTableShape(1)->createRow();

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = $oPhpPresentationRead->getSlide(0)->getShapeCollection();
        self::assertInstanceOf(Table::class, $arrayShape[0]);
        self::assertFalse($arrayShape[0]->isFirstRow());
        self::assertFalse($arrayShape[0]->isBandRow());
        self::assertInstanceOf(Table::class, $arrayShape[1]);
        self::assertTrue($arrayShape[1]->isFirstRow());
        self::assertTrue($arrayShape[1]->isBandRow());
    }

    /**
     * @return array<array{string, string}>
     */
    public static function dataProviderFields(): array
    {
        return [
            [Placeholder::PH_TYPE_SLIDENUM, 'slidenum'],
            [Placeholder::PH_TYPE_DATETIME, 'datetime'],
        ];
    }

    /**
     * @dataProvider dataProviderFields
     */
    #[DataProvider('dataProviderFields')]
    public function testFieldKeepsItsStyle(string $placeholder, string $type): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Page')
            ->getFont()
            ->setName('Georgia')
            ->setSize(18)
            ->setBold(true)
            ->setColor(new Color('FFCC0000'));
        $oShape->setPlaceHolder(new Placeholder($placeholder));

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);

        // The writer turns this shape into `a:fld` rather than a run, so the fixture is the file
        // itself: what the reader has to cope with is exactly what the writer produces.
        $oZip = new ZipArchive();
        $oZip->open($file);
        self::assertStringContainsString(
            '<a:fld id=',
            (string) $oZip->getFromName('ppt/slides/slide1.xml')
        );
        self::assertStringContainsString(
            'type="' . $type . '"',
            (string) $oZip->getFromName('ppt/slides/slide1.xml')
        );
        $oZip->close();

        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertTrue($arrayShape[0]->isPlaceholder());
        self::assertEquals($placeholder, $arrayShape[0]->getPlaceholder()->getType());

        // The text of a field is only the stand-in the writer put there; the styling is the shape's
        $arrayElements = $arrayShape[0]->getParagraph(0)->getRichTextElements();
        self::assertCount(1, $arrayElements);
        self::assertEquals('Georgia', $arrayElements[0]->getFont()->getName());
        self::assertEquals(18, $arrayElements[0]->getFont()->getSize());
        self::assertTrue($arrayElements[0]->getFont()->isBold());
        self::assertEquals('FFCC0000', $arrayElements[0]->getFont()->getColor()->getARGB());
    }

    public function testFieldIsReadBack(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oParagraph = $oPhpPresentation->getActiveSlide()->createRichTextShape()->getActiveParagraph();
        $oParagraph->createTextRun('page ');
        $oParagraph->createField(Field::TYPE_SLIDENUM, '<nr.>')->getFont()->setBold(true);
        $oParagraph->createTextRun(' of ');
        $oParagraph->createField(Field::TYPE_SLIDECOUNT, '12');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        $arrayElements = $arrayShape[0]->getActiveParagraph()->getRichTextElements();
        self::assertCount(4, $arrayElements);

        // what a field says is a stand-in for what it means, so the kind of field has to come back
        self::assertInstanceOf(Field::class, $arrayElements[1]);
        self::assertEquals(Field::TYPE_SLIDENUM, $arrayElements[1]->getType());
        self::assertEquals('<nr.>', $arrayElements[1]->getText());
        self::assertTrue($arrayElements[1]->getFont()->isBold());

        self::assertInstanceOf(Field::class, $arrayElements[3]);
        self::assertEquals(Field::TYPE_SLIDECOUNT, $arrayElements[3]->getType());

        // and the runs around it stay runs
        self::assertNotInstanceOf(Field::class, $arrayElements[0]);
        self::assertEquals('page ', $arrayElements[0]->getText());
    }

    /**
     * @dataProvider dataProviderShadowAlpha
     */
    #[DataProvider('dataProviderShadowAlpha')]
    public function testShapeShadowSurvivesTheRoundTrip(string $argb, int $alpha, string $expectedARGB): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Sample');
        $oShape->getShadow()
            ->setVisible(true)
            ->setAlignment(Shadow::SHADOW_BOTTOM_RIGHT)
            ->setColor(new Color($argb))
            ->setAlpha($alpha);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);

        // the alignment was read off `a:effectLst` rather than off the shadow inside it, and the
        // colour is written as `a:srgbClr` while only `a:prstClr` was read back
        $oShadow = $arrayShape[0]->getShadow();
        self::assertEquals(Shadow::SHADOW_BOTTOM_RIGHT, $oShadow->getAlignment());
        self::assertEquals($expectedARGB, $oShadow->getColor()->getARGB());
        self::assertEquals($alpha, $oShadow->getAlpha());
    }

    /**
     * A shadow states how see-through it is twice: on its own `alpha`, and in the two characters in
     * front of its colour. The file has room for one of them, `a:alpha` inside `a:srgbClr`, and it
     * is the shadow's that is written there -- so that is the one both come back carrying.
     *
     * @return array<array{string, int, string}>
     */
    public static function dataProviderShadowAlpha(): array
    {
        return [
            // an opaque colour on a shadow that is not: the colour gives way to what the file says
            ['FF00FF00', 40, '6600FF00'],
            // a colour see-through on its own account, on a shadow that is not
            ['80FF0000', 100, 'FFFF0000'],
            // the two agreeing, which is the only case where neither has to give way
            ['66FF0000', 40, '66FF0000'],
        ];
    }

    public function testShapeWrapSurvivesTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Sample');
        $oShape->setWrap(RichText::WRAP_NONE);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);

        // `wrap` is written on `a:bodyPr` beside the insets, and nothing read it back
        self::assertEquals(RichText::WRAP_NONE, $arrayShape[0]->getWrap());
    }

    public function testTextInsetsSurviveTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oShape = $oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Sample');
        $oShape->setInsetLeft(12.0)->setInsetTop(6.0)->setInsetRight(24.0)->setInsetBottom(3.0);

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);

        // an inset is EMU in the file and pixels in the model, so it comes back as the number that
        // went in and not as that number times the 9525 EMU of a pixel
        self::assertEquals(12.0, $arrayShape[0]->getInsetLeft());
        self::assertEquals(6.0, $arrayShape[0]->getInsetTop());
        self::assertEquals(24.0, $arrayShape[0]->getInsetRight());
        self::assertEquals(3.0, $arrayShape[0]->getInsetBottom());
    }

    public function testDefaultTextInsetsAreNotChangedByTheRoundTrip(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('Sample');

        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        (new PowerPoint2007Writer($oPhpPresentation))->save($file);
        $oPhpPresentationRead = (new PowerPoint2007())->load($file);
        unlink($file);

        $arrayShape = array_values((array) $oPhpPresentationRead->getActiveSlide()->getShapeCollection());
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertEquals(9.6, $arrayShape[0]->getInsetLeft());
        self::assertEquals(4.8, $arrayShape[0]->getInsetTop());
        self::assertEquals(9.6, $arrayShape[0]->getInsetRight());
        self::assertEquals(4.8, $arrayShape[0]->getInsetBottom());
    }
}
