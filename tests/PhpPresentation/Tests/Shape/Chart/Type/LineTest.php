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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart\Type;

use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Line element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Type\Line
 */
class LineTest extends TestCase
{
    public function testData(): void
    {
        $object = new Line();

        $this->assertIsArray($object->getSeries());
        $this->assertEmpty($object->getSeries());

        $array = [
            new Series(),
            new Series(),
        ];

        $this->assertInstanceOf(Line::class, $object->setSeries());
        $this->assertEmpty($object->getSeries());
        $this->assertInstanceOf(Line::class, $object->setSeries($array));
        $this->assertCount(count($array), $object->getSeries());
    }

    public function testSeries(): void
    {
        $object = new Line();

        $this->assertInstanceOf(Line::class, $object->addSeries(new Series()));
        $this->assertCount(1, $object->getSeries());
    }

    public function testSmooth(): void
    {
        $object = new Line();

        $this->assertFalse($object->isSmooth());
        $this->assertInstanceOf(Line::class, $object->setIsSmooth(true));
        $this->assertTrue($object->isSmooth());
    }

    public function testHashCode(): void
    {
        $series = new Series();

        $object = new Line();
        $object->addSeries($series);

        $this->assertEquals(
            md5(md5($object->isSmooth() ? '1' : '0') . $series->getHashCode() . get_class($object)),
            $object->getHashCode()
        );
    }
}
