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
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Scatter element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter
 */
class AbstractTest extends TestCase
{
    public function testAxis(): void
    {
        $object = new Scatter();

        $this->assertTrue($object->hasAxisX());
        $this->assertTrue($object->hasAxisY());
    }

    public function testHashIndex(): void
    {
        $object = new Scatter();
        $value = mt_rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testSeries(): void
    {
        /** @var AbstractType $stub */
        $stub = $this->getMockForAbstractClass(AbstractType::class);
        $this->assertEmpty($stub->getSeries());
        $this->assertIsArray($stub->getSeries());

        $arraySeries = [
            new Series(),
            new Series(),
        ];
        $this->assertInstanceOf(AbstractType::class, $stub->setSeries($arraySeries));
        $this->assertIsArray($stub->getSeries());
        $this->assertCount(count($arraySeries), $stub->getSeries());
    }

    public function testClone(): void
    {
        $arraySeries = [
            new Series(),
            new Series(),
            new Series(),
            new Series(),
        ];

        /** @var AbstractType $stub */
        $stub = $this->getMockForAbstractClass(AbstractType::class);
        $stub->setSeries($arraySeries);
        $clone = clone $stub;

        $this->assertInstanceOf(AbstractType::class, $clone);
        $this->assertIsArray($stub->getSeries());
        $this->assertCount(count($arraySeries), $stub->getSeries());
    }
}
