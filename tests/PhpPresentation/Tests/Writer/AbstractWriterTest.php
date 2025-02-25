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

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Tests\Writer\AbstractWriter as TestAbstractWriter;
use PhpOffice\PhpPresentation\Writer\AbstractWriter;
use PhpOffice\PhpPresentation\Writer\PDF\PDFWriterInterface;
use PHPUnit\Framework\TestCase;

require 'AbstractWriter.php';

/**
 * Test class for AbstractWriter.
 *
 * @coversDefaultClass \AbstractWriter
 *
 * @SuppressWarnings(PHPMD.UnusedFormalParameter)
 */
class AbstractWriterTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractWriter $mockWriter */
            $mockWriter = $this->getMockForAbstractClass(AbstractWriter::class);
        } else {
            /** @var AbstractWriter $mockWriter */
            $mockWriter = new class() extends AbstractWriter {
            };
        }

        self::assertNull($mockWriter->getPDFAdapter());
        self::assertNull($mockWriter->getZipAdapter());
    }

    public function testPDFAdapter(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractWriter $mockWriter */
            $mockWriter = $this->getMockForAbstractClass(AbstractWriter::class);
        } else {
            /** @var AbstractWriter $mockWriter */
            $mockWriter = new class() extends AbstractWriter {
            };
        }
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var PDFWriterInterface $mockPdfAdapter */
            $mockPdfAdapter = $this->getMockForAbstractClass(PDFWriterInterface::class);
        } else {
            /** @var PDFWriterInterface $mockPdfAdapter */
            $mockPdfAdapter = new class() implements PDFWriterInterface {
                public function save(string $filename): void
                {
                }

                public function setPhpPresentation(?PhpPresentation $pPhpPresentation = null)
                {
                    return $this;
                }
            };
        }

        self::assertNull($mockWriter->getPDFAdapter());
        self::assertInstanceOf(AbstractWriter::class, $mockWriter->setPDFAdapter($mockPdfAdapter));
        self::assertInstanceOf(PDFWriterInterface::class, $mockWriter->getPDFAdapter());
    }

    public function testZipAdapter(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractWriter $mockWriter */
            $mockWriter = $this->getMockForAbstractClass(AbstractWriter::class);
        } else {
            /** @var AbstractWriter $mockWriter */
            $mockWriter = new class() extends AbstractWriter {
            };
        }
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var ZipInterface $mockZip */
            $mockZip = $this->getMockForAbstractClass(ZipInterface::class);
        } else {
            /** @var ZipInterface $mockZip */
            $mockZip = new class() implements ZipInterface {
                public function open($filename)
                {
                    return $this;
                }

                public function close()
                {
                    return $this;
                }

                public function addFromString(string $localname, string $contents, bool $withCompression = true)
                {
                    return $this;
                }
            };
        }

        self::assertNull($mockWriter->getZipAdapter());
        self::assertInstanceOf(AbstractWriter::class, $mockWriter->setZipAdapter($mockZip));
        self::assertInstanceOf(ZipInterface::class, $mockWriter->getZipAdapter());
    }

    /**
     * Test all drawings method.
     */
    public function testAllDrawingsIncludesMasterSlides(): void
    {
        $presentation = new PhpPresentation();

        $activeSlide = $presentation->getActiveSlide();
        $activeSlide->createDrawingShape();

        $masterSlides = $presentation->getAllMasterSlides();
        $masterSlide = $masterSlides[0];
        $masterSlide->createDrawingShape();

        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var TestAbstractWriter $writer */
            $writer = $this->getMockForAbstractClass(TestAbstractWriter::class);
        } else {
            /** @var TestAbstractWriter $writer */
            $writer = new class() extends TestAbstractWriter {
            };
        }
        $writer->setPhpPresentation($presentation);

        $drawings = $writer->allDrawings();
        self::assertCount(2, $drawings, 'Number of drawings should equal two: one from normal slide and one from master slide');
    }
}
