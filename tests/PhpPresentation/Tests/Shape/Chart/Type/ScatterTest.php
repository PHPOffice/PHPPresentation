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
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart\Type;

use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Scatter element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter
 */
class ScatterTest extends TestCase
{
    public function testData(): void
    {
        $object = new Scatter();

        $this->assertIsArray($object->getSeries());
        $this->assertEmpty($object->getSeries());

        $array = [
            new Series(),
            new Series(),
        ];

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->setSeries());
        $this->assertEmpty($object->getSeries());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->setSeries($array));
        $this->assertCount(count($array), $object->getSeries());
    }

    public function testSerties(): void
    {
        $object = new Scatter();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->addSeries(new Series()));
        $this->assertCount(1, $object->getSeries());
    }

    public function testHashCode(): void
    {
        $oSeries = new Series();

        $object = new Scatter();
        $object->addSeries($oSeries);

        $this->assertEquals(md5($oSeries->getHashCode() . get_class($object)), $object->getHashCode());
    }
}
