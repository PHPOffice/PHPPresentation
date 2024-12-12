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

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Borders;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PhpPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\PhpPresentation
 */
class BordersTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        $object = new Borders();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getBottom());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getLeft());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getRight());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getTop());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getDiagonalDown());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Border', $object->getDiagonalUp());
        self::assertEquals(Border::LINE_NONE, $object->getDiagonalDown()->getLineStyle());
        self::assertEquals(Border::LINE_NONE, $object->getDiagonalUp()->getLineStyle());
    }

    /**
     * Test get/set hash index.
     */
    public function testSetGetHashIndex(): void
    {
        $object = new Borders();
        $value = mt_rand(1, 100);
        $object->setHashIndex($value);
        self::assertEquals($value, $object->getHashIndex());
    }

    /**
     * Test get/set hash code.
     */
    public function testGetHashCode(): void
    {
        $object = new Borders();
        self::assertEquals(
            md5(
                $object->getLeft()->getHashCode() .
                $object->getRight()->getHashCode() .
                $object->getTop()->getHashCode() .
                $object->getBottom()->getHashCode() .
                $object->getDiagonalUp()->getHashCode() .
                $object->getDiagonalDown()->getHashCode() .
                get_class($object)
            ),
            $object->getHashCode()
        );
    }
}
