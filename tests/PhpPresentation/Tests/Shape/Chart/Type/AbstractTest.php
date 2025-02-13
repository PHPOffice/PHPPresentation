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

        self::assertTrue($object->hasAxisX());
        self::assertTrue($object->hasAxisY());
    }

    public function testHashIndex(): void
    {
        $object = new Scatter();
        $value = mt_rand(1, 100);

        self::assertEmpty($object->getHashIndex());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->setHashIndex($value));
        self::assertEquals($value, $object->getHashIndex());
    }

    public function testSeries(): void
    {
        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractType $stub */
            $stub = $this->getMockForAbstractClass(AbstractType::class);
        } else {
            /** @var AbstractType $stub */
            $stub = new class() extends AbstractType {
                public function getHashCode(): string
                {
                    return '';
                }
            };
        }
        self::assertEmpty($stub->getSeries());
        self::assertIsArray($stub->getSeries());

        $arraySeries = [
            new Series(),
            new Series(),
        ];
        self::assertInstanceOf(AbstractType::class, $stub->setSeries($arraySeries));
        self::assertIsArray($stub->getSeries());
        self::assertCount(count($arraySeries), $stub->getSeries());
    }

    public function testClone(): void
    {
        $arraySeries = [
            new Series(),
            new Series(),
            new Series(),
            new Series(),
        ];

        if (method_exists($this, 'getMockForAbstractClass')) {
            /** @var AbstractType $stub */
            $stub = $this->getMockForAbstractClass(AbstractType::class);
        } else {
            /** @var AbstractType $stub */
            $stub = new class() extends AbstractType {
                public function getHashCode(): string
                {
                    return '';
                }
            };
        }
        $stub->setSeries($arraySeries);
        $clone = clone $stub;

        self::assertInstanceOf(AbstractType::class, $clone);
        self::assertIsArray($stub->getSeries());
        self::assertCount(count($arraySeries), $stub->getSeries());
    }
}
