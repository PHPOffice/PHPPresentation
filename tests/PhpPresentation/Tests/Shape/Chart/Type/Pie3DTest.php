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
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Pie3D element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D
 */
class Pie3DTest extends TestCase
{
    public function testData(): void
    {
        $object = new Pie3D();

        $this->assertIsArray($object->getSeries());
        $this->assertEmpty($object->getSeries());

        $array = [
            new Series(),
            new Series(),
        ];

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie3D', $object->setSeries());
        $this->assertEmpty($object->getSeries());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie3D', $object->setSeries($array));
        $this->assertCount(count($array), $object->getSeries());
    }

    public function testSeries(): void
    {
        $object = new Pie3D();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie3D', $object->addSeries(new Series()));
        $this->assertCount(1, $object->getSeries());
    }

    public function testExplosion(): void
    {
        $value = mt_rand(0, 100);
        $object = new Pie3D();

        $this->assertEquals(0, $object->getExplosion());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Pie3D', $object->setExplosion($value));
        $this->assertEquals($value, $object->getExplosion());
    }

    public function testHashCode(): void
    {
        $oSeries = new Series();

        $object = new Pie3D();
        $object->addSeries($oSeries);

        $this->assertEquals(md5($oSeries->getHashCode() . get_class($object)), $object->getHashCode());
    }
}
