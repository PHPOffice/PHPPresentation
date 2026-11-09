<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code.
 *
 * For the full list of contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Reader\PowerPoint2007;

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PHPUnit\Framework\TestCase;

/**
 * The geometry a picture is drawn in is read back from the file it was written to.
 *
 * A picture cropped to a shape is written with the preset of that shape, and a picture that was
 * cropped to a shape and read back is a picture cropped to a shape: the geometry is a part of the
 * picture rather than a part of the drawing that surrounds it.
 */
class DrawingShapeGeometryTest extends TestCase
{
    /**
     * @var null|string
     */
    private $filePath;

    protected function tearDown(): void
    {
        if (null !== $this->filePath) {
            @unlink($this->filePath);
        }
        $this->filePath = null;
    }

    public function testReadsTheGeometryAPictureIsDrawnIn(): void
    {
        $presentation = new PhpPresentation();
        $shape = $presentation->getActiveSlide()->createDrawingShape();
        $shape->setHeight(10)->setWidth(10);
        $shape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
        $shape->setGeometry(AutoShape::TYPE_OVAL);

        $read = $this->written($presentation);
        $shapes = $read->getActiveSlide()->getShapeCollection();

        self::assertCount(1, $shapes);
        self::assertInstanceOf(AbstractDrawingAdapter::class, $shapes[0]);
        self::assertSame(AutoShape::TYPE_OVAL, $shapes[0]->getGeometry());
    }

    public function testReadsAGeometryOfRectForAPictureThatWasGivenNone(): void
    {
        $presentation = new PhpPresentation();
        $shape = $presentation->getActiveSlide()->createDrawingShape();
        $shape->setHeight(10)->setWidth(10);
        $shape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');

        $read = $this->written($presentation);
        $shapes = $read->getActiveSlide()->getShapeCollection();

        self::assertCount(1, $shapes);
        self::assertInstanceOf(AbstractDrawingAdapter::class, $shapes[0]);
        self::assertSame('rect', $shapes[0]->getGeometry());
    }

    /**
     * Write the presentation to a file and read it back.
     */
    private function written(PhpPresentation $presentation): PhpPresentation
    {
        $this->filePath = tempnam(sys_get_temp_dir(), 'PhpPresentation');

        IOFactory::createWriter($presentation)->save($this->filePath);

        return IOFactory::createReader('PowerPoint2007')->load($this->filePath);
    }
}
