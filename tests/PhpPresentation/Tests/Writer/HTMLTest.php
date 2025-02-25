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
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\HTML;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\HTML.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\HTML
 */
class HTMLTest extends PhpPresentationTestCase
{
    protected $writerName = 'HTML';

    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $this->oPresentation->getActiveSlide()->createDrawingShape();
        $object = new HTML($this->oPresentation);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\HashTable', $object->getDrawingHashTable());
    }

    /**
     * Test save.
     */
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

        $object = new HTML($this->oPresentation);
        $object->save($filename);

        self::assertFileExists($filename);

        unlink($filename);
    }

    /**
     * Test get PhpPresentation exception.
     */
    public function testSaveEmpty(): void
    {
        $this->expectException(InvalidParameterException::class);
        $this->expectExceptionMessage('The parameter pFilename can\'t have the value ""');

        $object = new HTML();
        $object->save('');
    }
}
