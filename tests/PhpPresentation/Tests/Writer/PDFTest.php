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

use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\Exception\WriterPDFAdapterNotDefinedException;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\PDF;
use PhpOffice\PhpPresentation\Writer\PDF\DomPDF;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\PDF.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\PDF
 */
class PDFTest extends PhpPresentationTestCase
{
    protected $writerName = 'PDF';

    public function testConstruct(): void
    {
        $this->oPresentation->getActiveSlide()->createDrawingShape();
        $object = new PDF($this->oPresentation);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\HashTable', $object->getDrawingHashTable());
    }

    public function testSave(): void
    {
        $filename = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        $imageFile = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png';

        $slide = $this->oPresentation->getActiveSlide();
        $slide->createRichTextShape();
        $slide->createLineShape(10, 10, 10, 10);
        $slide->createChartShape()->getPlotArea()->setType(new Bar3D());
        $slide->createDrawingShape()->setName('Drawing')->setPath($imageFile);
        $slide->createTableShape()->createRow();

        $object = new PDF($this->oPresentation);
        $object->setPDFAdapter(new DomPDF());
        $object->save($filename);

        self::assertFileExists($filename);

        unlink($filename);
    }

    public function testSaveEmpty(): void
    {
        $this->expectException(InvalidParameterException::class);
        $this->expectExceptionMessage('The parameter pFilename can\'t have the value ""');

        $object = new PDF();
        $object->save('');
    }

    public function testSaveNoPDFAdapter(): void
    {
        $this->expectException(WriterPDFAdapterNotDefinedException::class);
        $this->expectExceptionMessage('The PDF Adapter has not been defined');

        $object = new PDF();
        $object->save('file.pdf');
    }
}
