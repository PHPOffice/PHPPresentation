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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Reader\Keynote as KeynoteReader;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Writer\Keynote as KeynoteWriter;
use PHPUnit\Framework\TestCase;

/**
 * Round-trip test for the Keynote reader and writer.
 */
class KeynoteRoundtripTest extends TestCase
{
    public function testRoundtripPreservesContent(): void
    {
        $oOriginal = new PhpPresentation();
        $oOriginal->removeSlideByIndex();

        $oSlide1 = $oOriginal->createSlide();
        $oShape1 = $oSlide1->createRichTextShape();
        $oShape1->setOffsetX(100);
        $oShape1->setOffsetY(200);
        $oShape1->setWidth(300);
        $oShape1->setHeight(150);
        $oParagraph1 = $oShape1->getActiveParagraph();
        $oParagraph1->createTextRun('First ');
        $oParagraph1->createTextRun('run');
        $oParagraph2 = $oShape1->createParagraph();
        $oParagraph2->createTextRun('Second paragraph');
        $oSlide1->getNote()->createRichTextShape()->createTextRun('Note text');

        $oSlide2 = $oOriginal->createSlide();
        $oSlide2->createRichTextShape()->createTextRun('Slide two');

        $sFilename = tempnam(sys_get_temp_dir(), 'kny') . '.key';
        (new KeynoteWriter($oOriginal))->save($sFilename);

        $oResult = (new KeynoteReader())->load($sFilename);
        unlink($sFilename);

        self::assertCount(2, $oResult->getAllSlides());

        $aShapes = $oResult->getSlide(0)->getShapeCollection();
        self::assertCount(1, $aShapes);
        $oResultShape = $aShapes[0];
        self::assertInstanceOf(RichText::class, $oResultShape);
        self::assertSame(100, $oResultShape->getOffsetX());
        self::assertSame(200, $oResultShape->getOffsetY());
        self::assertSame(300, $oResultShape->getWidth());
        self::assertSame(150, $oResultShape->getHeight());

        $aParagraphs = $oResultShape->getParagraphs();
        self::assertCount(2, $aParagraphs);
        self::assertSame('First run', $aParagraphs[0]->getPlainText());
        self::assertSame('Second paragraph', $aParagraphs[1]->getPlainText());

        $aNoteShapes = $oResult->getSlide(0)->getNote()->getShapeCollection();
        self::assertCount(1, $aNoteShapes);
        $oNoteShape = $aNoteShapes[0];
        self::assertInstanceOf(RichText::class, $oNoteShape);
        $sNoteText = $oNoteShape->getPlainText();
        self::assertSame('Note text', $sNoteText);

        self::assertCount(1, $oResult->getSlide(1)->getShapeCollection());
    }

    public function testRoundtripEmptyPresentation(): void
    {
        $oOriginal = new PhpPresentation();
        $oOriginal->removeSlideByIndex();

        $sFilename = tempnam(sys_get_temp_dir(), 'kny') . '.key';
        (new KeynoteWriter($oOriginal))->save($sFilename);

        $oResult = (new KeynoteReader())->load($sFilename);
        unlink($sFilename);

        self::assertCount(0, $oResult->getAllSlides());
    }

    /**
     * Text runs carrying non-ASCII characters and XML metacharacters must survive
     * a write/read cycle unchanged (proper escaping in the writer, unescaping in
     * the reader).
     */
    public function testRoundtripPreservesUnicodeAndSpecialCharacters(): void
    {
        $sText = 'Unicode: aeiou 日本語 emoji 😀 - markup <tag> & "quotes" \'apostrophe\'';

        $oOriginal = new PhpPresentation();
        $oOriginal->removeSlideByIndex();
        $oOriginal->createSlide()->createRichTextShape()->createTextRun($sText);

        $sFilename = tempnam(sys_get_temp_dir(), 'kny') . '.key';
        (new KeynoteWriter($oOriginal))->save($sFilename);

        $oResult = (new KeynoteReader())->load($sFilename);
        unlink($sFilename);

        $aShapes = $oResult->getSlide(0)->getShapeCollection();
        self::assertCount(1, $aShapes);
        $oResultShape = $aShapes[0];
        self::assertInstanceOf(RichText::class, $oResultShape);
        self::assertSame($sText, $oResultShape->getPlainText());
    }

    /**
     * Every text placeholder of a slide is written and read back in order.
     */
    public function testRoundtripPreservesMultipleShapesPerSlide(): void
    {
        $oOriginal = new PhpPresentation();
        $oOriginal->removeSlideByIndex();
        $oSlide = $oOriginal->createSlide();
        $oSlide->createRichTextShape()->createTextRun('Shape one');
        $oSlide->createRichTextShape()->createTextRun('Shape two');

        $sFilename = tempnam(sys_get_temp_dir(), 'kny') . '.key';
        (new KeynoteWriter($oOriginal))->save($sFilename);

        $oResult = (new KeynoteReader())->load($sFilename);
        unlink($sFilename);

        $aShapes = $oResult->getSlide(0)->getShapeCollection();
        self::assertCount(2, $aShapes);
        self::assertInstanceOf(RichText::class, $aShapes[0]);
        self::assertInstanceOf(RichText::class, $aShapes[1]);
        self::assertSame('Shape one', $aShapes[0]->getPlainText());
        self::assertSame('Shape two', $aShapes[1]->getPlainText());
    }

    /**
     * Content that the basic format cannot represent is skipped silently: a
     * non-text shape produces no output, and line breaks inside a text run are
     * dropped while the surrounding runs are preserved.
     */
    public function testRoundtripSkipsNonRepresentableContent(): void
    {
        $oOriginal = new PhpPresentation();
        $oOriginal->removeSlideByIndex();
        $oSlide = $oOriginal->createSlide();

        // A line is not a text placeholder and must not be written.
        $oSlide->createLineShape(10, 10, 100, 100);

        $oShape = $oSlide->createRichTextShape();
        $oParagraph = $oShape->getActiveParagraph();
        $oParagraph->createTextRun('before');
        $oParagraph->createBreak();
        $oParagraph->createTextRun('after');

        $sFilename = tempnam(sys_get_temp_dir(), 'kny') . '.key';
        (new KeynoteWriter($oOriginal))->save($sFilename);

        $oResult = (new KeynoteReader())->load($sFilename);
        unlink($sFilename);

        // Only the text placeholder survives; the line shape is skipped.
        $aShapes = $oResult->getSlide(0)->getShapeCollection();
        self::assertCount(1, $aShapes);
        $oResultShape = $aShapes[0];
        self::assertInstanceOf(RichText::class, $oResultShape);
        // The break is dropped, so the two runs are concatenated.
        self::assertSame('beforeafter', $oResultShape->getPlainText());
    }
}
