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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Shape\Chart\View3D;
use PHPUnit\Framework\TestCase;

/**
 * Test class for View3D element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\View3D
 */
class View3DTest extends TestCase
{
    public function testDepthPercent(): void
    {
        $object = new View3D();
        $value = mt_rand(20, 20000);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setDepthPercent());
        self::assertEquals(100, $object->getDepthPercent());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setDepthPercent($value));
        self::assertEquals($value, $object->getDepthPercent());
    }

    public function testHashIndex(): void
    {
        $object = new View3D();
        $value = mt_rand(1, 100);

        self::assertEmpty($object->getHashIndex());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setHashIndex($value));
        self::assertEquals($value, $object->getHashIndex());
    }

    public function testHeightPercent(): void
    {
        $object = new View3D();
        $value = mt_rand(5, 500);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setHeightPercent());
        self::assertEquals(100, $object->getHeightPercent());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setHeightPercent($value));
        self::assertEquals($value, $object->getHeightPercent());
    }

    public function testPerspective(): void
    {
        $object = new View3D();
        $value = mt_rand(0, 100);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setPerspective());
        self::assertEquals(30, $object->getPerspective());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setPerspective($value));
        self::assertEquals($value, $object->getPerspective());
    }

    public function testRightAngleAxes(): void
    {
        $object = new View3D();

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setRightAngleAxes());
        self::assertTrue($object->hasRightAngleAxes());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setRightAngleAxes(true));
        self::assertTrue($object->hasRightAngleAxes());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setRightAngleAxes(false));
        self::assertFalse($object->hasRightAngleAxes());
    }

    public function testRotationX(): void
    {
        $object = new View3D();
        $value = mt_rand(-90, 90);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setRotationX());
        self::assertEquals(0, $object->getRotationX());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setRotationX($value));
        self::assertEquals($value, $object->getRotationX());
    }

    public function testRotationY(): void
    {
        $object = new View3D();
        $value = mt_rand(-90, 90);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setRotationY());
        self::assertEquals(0, $object->getRotationY());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\View3D', $object->setRotationY($value));
        self::assertEquals($value, $object->getRotationY());
    }
}
