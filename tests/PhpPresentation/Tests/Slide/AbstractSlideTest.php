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

namespace PhpOffice\PhpPresentation\Tests\Slide;

use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide\AbstractSlide;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Table element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\AbstractGraphic
 */
class AbstractSlideTest extends TestCase
{
    public function testCollection(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractSlide $object */
            $object = $this->getMockForAbstractClass(AbstractSlide::class);
        } else {
            /** @var AbstractSlide $object */
            $object = new class() extends AbstractSlide {
            };
        }
        $array = [];
        self::assertInstanceOf(AbstractSlide::class, $object->setShapeCollection($array));
        self::assertIsArray($object->getShapeCollection());
        self::assertCount(count($array), $object->getShapeCollection());

        $array = [
            new RichText(),
            new RichText(),
            new RichText(),
        ];
        self::assertInstanceOf(AbstractSlide::class, $object->setShapeCollection($array));
        self::assertIsArray($object->getShapeCollection());
        self::assertCount(count($array), $object->getShapeCollection());
    }

    public function testAdd(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractSlide $object */
            $object = $this->getMockForAbstractClass(AbstractSlide::class);
        } else {
            /** @var AbstractSlide $object */
            $object = new class() extends AbstractSlide {
            };
        }

        self::assertInstanceOf(AutoShape::class, $object->createAutoShape());
        self::assertInstanceOf(Chart::class, $object->createChartShape());
        self::assertInstanceOf(Drawing\File::class, $object->createDrawingShape());
        self::assertInstanceOf(Line::class, $object->createLineShape(10, 10, 10, 10));
        self::assertInstanceOf(RichText::class, $object->createRichTextShape());
        self::assertInstanceOf(Table::class, $object->createTableShape());
        self::assertCount(6, $object->getShapeCollection());
    }

    public function testSearchShapes(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractSlide $object */
            $object = $this->getMockForAbstractClass(AbstractSlide::class);
        } else {
            /** @var AbstractSlide $object */
            $object = new class() extends AbstractSlide {
            };
        }

        $array = [
            (new RichText())->setName('AAA'),
            (new Table())->setName('BBB'),
            (new Chart())->setName('AAA'),
        ];
        self::assertInstanceOf(AbstractSlide::class, $object->setShapeCollection($array));

        // Search by Name
        $result = $object->searchShapes('AAA', null);
        self::assertIsArray($result);
        self::assertCount(2, $result);
        self::assertInstanceOf(RichText::class, $result[0]);
        self::assertInstanceOf(Chart::class, $result[1]);

        // Search by Name && Type
        $result = $object->searchShapes('AAA', Chart::class);
        self::assertIsArray($result);
        self::assertCount(1, $result);
        self::assertInstanceOf(Chart::class, $result[0]);

        // Search by Type
        $result = $object->searchShapes(null, Table::class);
        self::assertIsArray($result);
        self::assertCount(1, $result);
        self::assertInstanceOf(Table::class, $result[0]);
    }
}
