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

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Iterator;
use PHPUnit\Framework\TestCase;

/**
 * Test class for IOFactory.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\IOFactory
 */
class IteratorTest extends TestCase
{
    public function testMethod(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->addSlide(new Slide());

        $object = new Iterator($oPhpPresentation);

        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide', $object->current());
        self::assertEquals(0, $object->key());
        $object->next();
        self::assertEquals(1, $object->key());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide', $object->current());
        self::assertTrue($object->valid());
        $object->next();
        self::assertFalse($object->valid());
        $object->rewind();
        self::assertEquals(0, $object->key());
    }
}
