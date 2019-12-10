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
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Slide;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Iterator;
use PHPUnit\Framework\TestCase;

/**
 * Test class for IOFactory
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\IOFactory
 */
class IteratorTest extends TestCase
{
    /**
     */
    public function testMethod()
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->addSlide(new Slide());

        $object = new Iterator($oPhpPresentation);

        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide', $object->current());
        static::assertEquals(0, $object->key());
        static::assertNull($object->next());
        static::assertEquals(1, $object->key());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Slide', $object->current());
        static::assertTrue($object->valid());
        static::assertNull($object->next());
        static::assertFalse($object->valid());
        static::assertNull($object->rewind());
        static::assertEquals(0, $object->key());
    }
}
