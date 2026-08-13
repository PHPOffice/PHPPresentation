<?php

declare(strict_types=1);

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PHPUnit\Framework\TestCase;
use ZipArchive;

class ContentTypesTest extends TestCase
{
    public function testSvgDefaultContentTypeIsWritten(): void
    {
        $presentation = new PhpPresentation();

        $pptxFile = tempnam(sys_get_temp_dir(), 'pptx_svg_test_');
        $writer = IOFactory::createWriter($presentation, 'PowerPoint2007');
        $writer->save($pptxFile);

        $zip = new ZipArchive();
        self::assertTrue(
            $zip->open($pptxFile),
            'Could not open generated PPTX as ZipArchive'
        );

        $contentTypesXml = $zip->getFromName('[Content_Types].xml');
        $zip->close();
        @unlink($pptxFile);

        self::assertIsString(
            $contentTypesXml,
            '[Content_Types].xml not found in archive'
        );

        self::assertStringContainsString(
            'Extension="svg"',
            $contentTypesXml,
            'SVG extension not registered in [Content_Types].xml'
        );

        self::assertStringContainsString(
            'image/svg+xml',
            $contentTypesXml,
            'SVG MIME type not registered correctly in [Content_Types].xml'
        );
    }
}
