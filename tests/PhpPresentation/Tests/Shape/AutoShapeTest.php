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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;
use PHPUnit\Framework\TestCase;
use ZipArchive;

class AutoShapeTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new AutoShape();

        self::assertEquals(AutoShape::TYPE_HEART, $object->getType());
        self::assertEquals('', $object->getText());
        self::assertInstanceOf(Outline::class, $object->getOutline());
        self::assertIsString($object->getHashCode());
    }

    public function testOutline(): void
    {
        /** @var Outline $mock */
        $mock = $this->getMockBuilder(Outline::class)->getMock();

        $object = new AutoShape();
        self::assertInstanceOf(Outline::class, $object->getOutline());
        self::assertInstanceOf(AutoShape::class, $object->setOutline($mock));
        self::assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testText(): void
    {
        $object = new AutoShape();

        self::assertEquals('', $object->getText());
        self::assertInstanceOf(AutoShape::class, $object->setText('Text'));
        self::assertEquals('Text', $object->getText());
    }

    public function testType(): void
    {
        $object = new AutoShape();

        self::assertEquals(AutoShape::TYPE_HEART, $object->getType());
        self::assertInstanceOf(AutoShape::class, $object->setType(AutoShape::TYPE_HEXAGON));
        self::assertEquals(AutoShape::TYPE_HEXAGON, $object->getType());
    }

    public function testPixelSetterComputesAdjAndAffectsHash(): void
    {
        $w = 200;  // px
        $h = 100;  // px
        $px1 = 5;  // softer radius
        $px2 = 10; // larger radius

        $s1 = (new AutoShape())
            ->setType(AutoShape::TYPE_ROUNDED_RECTANGLE)
            ->setWidth($w)->setHeight($h)
            ->setRoundRectCorner($px1);

        $s2 = (clone $s1)->setRoundRectCorner($px2);

        // adj expected: round(px / (min(w,h)/2) * 50000)
        $minHalf = (int) floor(min($w, $h) / 2); // 50
        $expectedAdj1 = (int) round($px1 / $minHalf * 50000); // 5/50 * 50000 = 5000
        $expectedAdj2 = (int) round($px2 / $minHalf * 50000); // 10/50 * 50000 = 10000

        self::assertSame($expectedAdj1, $s1->getRoundRectAdj());
        self::assertSame($expectedAdj2, $s2->getRoundRectAdj());

        // Hash must differ when radius differs
        self::assertNotSame($s1->getHashCode(), $s2->getHashCode());
    }

    public function testNoRadiusByDefaultIsNull(): void
    {
        $shape = new AutoShape();
        self::assertNull($shape->getRoundRectAdj());
    }

    public function testWriterEmitsAdjGuideForRoundRect(): void
    {
        $ppt = new PhpPresentation();
        $slide = $ppt->getActiveSlide();

        $width = 200;
        $height = 100;
        $padding = 5;
        $minHalf = (int) floor(min($width, $height) / 2);
        $expectedAdj = (int) round($padding / $minHalf * 50000); // 5000

        $shape = (new AutoShape())
            ->setType(AutoShape::TYPE_ROUNDED_RECTANGLE)
            ->setWidth($width)->setHeight($height)
            ->setRoundRectCorner($padding);

        // Give it a fill so it's an obvious shape
        $shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));
        $slide->addShape($shape);

        $tmpFile = tempnam(sys_get_temp_dir(), 'pptx_');
        $writer = new PowerPoint2007($ppt);
        $writer->save($tmpFile);

        // Open the pptx and read slide1.xml
        $zip = new ZipArchive();
        $this->assertTrue($zip->open($tmpFile) === true, 'Failed to open pptx zip');
        $xml = $zip->getFromName('ppt/slides/slide1.xml');
        $zip->close();
        @unlink($tmpFile);

        $this->assertIsString($xml);

        // Must contain roundRect geometry and the adj guide with expected value
        $this->assertStringContainsString('<a:prstGeom prst="roundRect">', $xml);

        // fmla="val N" (there is a space after 'val' in writer)
        $this->assertMatchesRegularExpression(
            sprintf('/<a:gd[^>]+name="adj"[^>]+fmla="val %d"/', $expectedAdj),
            $xml
        );
    }
}
