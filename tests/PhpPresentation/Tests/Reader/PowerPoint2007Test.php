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

namespace PhpOffice\PhpPresentation\Tests\Reader;

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Reader\PowerPoint2007;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PowerPoint2007 reader
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Reader\PowerPoint2007
 */
class PowerPoint2007Test extends TestCase
{
    /**
     * Test can read
     */
    public function testCanRead()
    {
        $object = new PowerPoint2007();

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt';
        $this->assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
        $this->assertFalse($object->canRead($file));

        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $this->assertTrue($object->canRead($file));
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Could not open  for reading! File does not exist.
     */
    public function testLoadFileNotExists()
    {
        $object = new PowerPoint2007();
        $object->load('');
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Invalid file format for PhpOffice\PhpPresentation\Reader\PowerPoint2007:
     */
    public function testLoadFileBadFormat()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.ppt';
        $object = new PowerPoint2007();
        $object->load($file);
    }

    /**
     * @expectedException \Exception
     * @expectedExceptionMessage Could not open  for reading! File does not exist.
     */
    public function testFileSupportsNotExists()
    {
        $object = new PowerPoint2007();
        $object->fileSupportsUnserializePhpPresentation('');
    }

    public function testLoadFile01()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        // Document Properties
        $this->assertEquals('PHPOffice', $oPhpPresentation->getDocumentProperties()->getCreator());
        $this->assertEquals('PHPPresentation Team', $oPhpPresentation->getDocumentProperties()->getLastModifiedBy());
        $this->assertEquals('Sample 02 Title', $oPhpPresentation->getDocumentProperties()->getTitle());
        $this->assertEquals('Sample 02 Subject', $oPhpPresentation->getDocumentProperties()->getSubject());
        $this->assertEquals('Sample 02 Description', $oPhpPresentation->getDocumentProperties()->getDescription());
        $this->assertEquals('office 2007 openxml libreoffice odt php', $oPhpPresentation->getDocumentProperties()->getKeywords());
        $this->assertEquals('Sample Category', $oPhpPresentation->getDocumentProperties()->getCategory());
        // Document Layout
        $this->assertEquals(DocumentLayout::LAYOUT_SCREEN_4X3, $oPhpPresentation->getLayout()->getDocumentLayout());
        $this->assertEquals(254, $oPhpPresentation->getLayout()->getCX(DocumentLayout::UNIT_MILLIMETER));
        $this->assertEquals(190.5, $oPhpPresentation->getLayout()->getCY(DocumentLayout::UNIT_MILLIMETER));

        // Slides
        $this->assertCount(4, $oPhpPresentation->getAllSlides());

        // Slide 1
        $oSlide1 = $oPhpPresentation->getSlide(0);
        $arrayShape = $oSlide1->getShapeCollection();
        $this->assertCount(2, $arrayShape);
        // Slide 1 : Shape 1
        $oShape = $arrayShape[0];
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\Gd', $oShape);
        $this->assertEquals('PHPPresentation logo', $oShape->getName());
        $this->assertEquals('PHPPresentation logo', $oShape->getDescription());
        $this->assertEquals(36, $oShape->getHeight());
        $this->assertEquals(10, $oShape->getOffsetX());
        $this->assertEquals(10, $oShape->getOffsetY());
        $this->assertTrue($oShape->getShadow()->isVisible());
        static::assertEquals(45, $oShape->getShadow()->getDirection());
        static::assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 1 : Shape 2
        $oShape = $arrayShape[1];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $oShape);
        static::assertEquals(200, $oShape->getHeight());
        static::assertEquals(600, $oShape->getWidth());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(400, $oShape->getOffsetY());
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oShape->getActiveParagraph()->getAlignment()->getHorizontal());
        $arrayParagraphs = $oShape->getParagraphs();
        static::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(3, $arrayRichText);
        // Slide 1 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Introduction to', $oRichText->getText());
        static::assertTrue($oRichText->getFont()->isBold());
        static::assertEquals(28, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 1 : Shape 2 : Paragraph 2
        $oRichText = $arrayRichText[1];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $oRichText);
        // Slide 1 : Shape 2 : Paragraph 3
        $oRichText = $arrayRichText[2];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('PHPPresentation', $oRichText->getText());
        static::assertTrue($oRichText->getFont()->isBold());
        static::assertEquals(60, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());

        // Slide 2
        $oSlide2 = $oPhpPresentation->getSlide(1);
        $arrayShape = $oSlide2->getShapeCollection();
        static::assertCount(3, $arrayShape);
        // Slide 2 : Shape 1
        $oShape = $arrayShape[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\Gd', $oShape);
        static::assertEquals('PHPPresentation logo', $oShape->getName());
        static::assertEquals('PHPPresentation logo', $oShape->getDescription());
        static::assertEquals(36, $oShape->getHeight());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(10, $oShape->getOffsetY());
        static::assertTrue($oShape->getShadow()->isVisible());
        static::assertEquals(45, $oShape->getShadow()->getDirection());
        static::assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 2 : Shape 2
        $oShape = $arrayShape[1];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $oShape);
        static::assertEquals(100, $oShape->getHeight());
        static::assertEquals(930, $oShape->getWidth());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(50, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        static::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        // Slide 2 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('What is PHPPresentation?', $oRichText->getText());
        static::assertTrue($oRichText->getFont()->isBold());
        static::assertEquals(48, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 2 : Shape 3
        $oShape = $arrayShape[2];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $oShape);
        static::assertEquals(600, $oShape->getHeight());
        static::assertEquals(930, $oShape->getWidth());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        static::assertCount(4, $arrayParagraphs);
        // Slide 2 : Shape 3 : Paragraph 1
        $oParagraph = $arrayParagraphs[0];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('A class library', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 2 : Shape 3 : Paragraph 2
        $oParagraph = $arrayParagraphs[1];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Written in PHP', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 2 : Shape 3 : Paragraph 3
        $oParagraph = $arrayParagraphs[2];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Representing a presentation', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 2 : Shape 3 : Paragraph 4
        $oParagraph = $arrayParagraphs[3];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Supports writing to different file formats', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());

        // Slide 3
        $oSlide2 = $oPhpPresentation->getSlide(2);
        $arrayShape = $oSlide2->getShapeCollection();
        static::assertCount(3, $arrayShape);
        // Slide 3 : Shape 1
        $oShape = $arrayShape[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\Gd', $oShape);
        static::assertEquals('PHPPresentation logo', $oShape->getName());
        static::assertEquals('PHPPresentation logo', $oShape->getDescription());
        static::assertEquals(36, $oShape->getHeight());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(10, $oShape->getOffsetY());
        static::assertTrue($oShape->getShadow()->isVisible());
        static::assertEquals(45, $oShape->getShadow()->getDirection());
        static::assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 3 : Shape 2
        $oShape = $arrayShape[1];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $oShape);
        static::assertEquals(100, $oShape->getHeight());
        static::assertEquals(930, $oShape->getWidth());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(50, $oShape->getOffsetY());
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oShape->getActiveParagraph()->getAlignment()->getHorizontal());
        $arrayParagraphs = $oShape->getParagraphs();
        static::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        // Slide 3 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('What\'s the point?', $oRichText->getText());
        static::assertTrue($oRichText->getFont()->isBold());
        static::assertEquals(48, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 3 : Shape 2
        $oShape = $arrayShape[2];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $oShape);
        static::assertEquals(600, $oShape->getHeight());
        static::assertEquals(930, $oShape->getWidth());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(130, $oShape->getOffsetY());
        $arrayParagraphs = $oShape->getParagraphs();
        static::assertCount(8, $arrayParagraphs);
        // Slide 3 : Shape 3 : Paragraph 1
        $oParagraph = $arrayParagraphs[0];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(0, $oParagraph->getAlignment()->getLevel());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Generate slide decks', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 3 : Shape 3 : Paragraph 2
        $oParagraph = $arrayParagraphs[1];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Represent business data', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 3 : Shape 3 : Paragraph 3
        $oParagraph = $arrayParagraphs[2];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Show a family slide show', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 3 : Shape 3 : Paragraph 4
        $oParagraph = $arrayParagraphs[3];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('...', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 3 : Shape 3 : Paragraph 5
        $oParagraph = $arrayParagraphs[4];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(25, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(0, $oParagraph->getAlignment()->getLevel());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Export these to different formats', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 3 : Shape 3 : Paragraph 6
        $oParagraph = $arrayParagraphs[5];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('PHPPresentation 2007', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 3 : Shape 3 : Paragraph 7
        $oParagraph = $arrayParagraphs[6];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Serialized', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());
        // Slide 3 : Shape 3 : Paragraph 8
        $oParagraph = $arrayParagraphs[7];
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oParagraph->getAlignment()->getHorizontal());
        static::assertEquals(75, $oParagraph->getAlignment()->getMarginLeft());
        static::assertEquals(-25, $oParagraph->getAlignment()->getIndent());
        static::assertEquals(1, $oParagraph->getAlignment()->getLevel());
        static::assertEquals(Bullet::TYPE_BULLET, $oParagraph->getBulletStyle()->getBulletType());
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('... (more to come) ...', $oRichText->getText());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oRichText->getFont()->getColor()->getARGB());

        // Slide 4
        $oSlide3 = $oPhpPresentation->getSlide(3);
        $arrayShape = $oSlide3->getShapeCollection();
        static::assertCount(3, $arrayShape);
        // Slide 4 : Shape 1
        $oShape = $arrayShape[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\Gd', $oShape);
        static::assertEquals('PHPPresentation logo', $oShape->getName());
        static::assertEquals('PHPPresentation logo', $oShape->getDescription());
        static::assertEquals(36, $oShape->getHeight());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(10, $oShape->getOffsetY());
        static::assertTrue($oShape->getShadow()->isVisible());
        static::assertEquals(45, $oShape->getShadow()->getDirection());
        static::assertEquals(10, $oShape->getShadow()->getDistance());
        // Slide 4 : Shape 2
        $oShape = $arrayShape[1];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $oShape);
        static::assertEquals(100, $oShape->getHeight());
        static::assertEquals(930, $oShape->getWidth());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(50, $oShape->getOffsetY());
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oShape->getActiveParagraph()->getAlignment()->getHorizontal());
        $arrayParagraphs = $oShape->getParagraphs();
        static::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(1, $arrayRichText);
        // Slide 4 : Shape 2 : Paragraph 1
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Need more info?', $oRichText->getText());
        static::assertTrue($oRichText->getFont()->isBold());
        static::assertEquals(48, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oShape->getActiveParagraph()->getFont()->getColor()->getARGB());
        // Slide 4 : Shape 3
        $oShape = $arrayShape[2];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText', $oShape);
        static::assertEquals(600, $oShape->getHeight());
        static::assertEquals(930, $oShape->getWidth());
        static::assertEquals(10, $oShape->getOffsetX());
        static::assertEquals(130, $oShape->getOffsetY());
        static::assertEquals(Alignment::HORIZONTAL_LEFT, $oShape->getActiveParagraph()->getAlignment()->getHorizontal());
        $arrayParagraphs = $oShape->getParagraphs();
        static::assertCount(1, $arrayParagraphs);
        $oParagraph = $arrayParagraphs[0];
        $arrayRichText = $oParagraph->getRichTextElements();
        static::assertCount(3, $arrayRichText);
        // Slide 4 : Shape 3 : Paragraph 1
        $oRichText = $arrayRichText[0];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('Check the project site on GitHub:', $oRichText->getText());
        static::assertFalse($oRichText->getFont()->isBold());
        static::assertEquals(36, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oShape->getActiveParagraph()->getFont()->getColor()->getARGB());
        // Slide 4 : Shape 3 : Paragraph 2
        $oRichText = $arrayRichText[1];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $oRichText);
        // Slide 4 : Shape 3 : Paragraph 3
        $oRichText = $arrayRichText[2];
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $oRichText);
        static::assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getText());
        static::assertFalse($oRichText->getFont()->isBold());
        static::assertEquals(32, $oRichText->getFont()->getSize());
        static::assertEquals('FF000000', $oShape->getActiveParagraph()->getFont()->getColor()->getARGB());
        static::assertTrue($oRichText->hasHyperlink());
        static::assertEquals('https://github.com/PHPOffice/PHPPresentation/', $oRichText->getHyperlink()->getUrl());
        static::assertEquals('PHPPresentation', $oRichText->getHyperlink()->getTooltip());
    }

    public function testMarkAsFinal()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        static::assertFalse($oPhpPresentation->getPresentationProperties()->isMarkedAsFinal());


        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_MarkAsFinal.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        static::assertTrue($oPhpPresentation->getPresentationProperties()->isMarkedAsFinal());
    }

    public function testZoom()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        static::assertEquals(1, $oPhpPresentation->getPresentationProperties()->getZoom());


        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/PPTX_Zoom.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);
        static::assertEquals(2.68, $oPhpPresentation->getPresentationProperties()->getZoom());
    }

    public function testSlideLayout()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Issue_00150.pptx';
        $object = new PowerPoint2007();
        $oPhpPresentation = $object->load($file);
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $oPhpPresentation);

        $masterSlides = $oPhpPresentation->getAllMasterSlides();
        static::assertCount(3, $masterSlides);
        static::assertCount(11, $masterSlides[0]->getAllSlideLayouts());
        static::assertCount(11, $masterSlides[1]->getAllSlideLayouts());
        static::assertCount(11, $masterSlides[2]->getAllSlideLayouts());
    }
}
