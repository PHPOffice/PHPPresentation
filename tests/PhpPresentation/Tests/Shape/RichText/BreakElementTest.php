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

namespace PhpOffice\PhpPresentation\Tests\Shape\RichText;

use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PHPUnit\Framework\TestCase;

/**
 * Test class for BreakElement element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\RichText\BreakElement
 */
class BreakElementTest extends TestCase
{
    /**
     * Test can read.
     */
    public function testText(): void
    {
        $object = new BreakElement();
        self::assertEquals("\r\n", $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->setText());
        self::assertEquals("\r\n", $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->setText('AAA'));
        self::assertEquals("\r\n", $object->getText());
    }

    public function testFont(): void
    {
        $object = new BreakElement();
        self::assertNull($object->getFont());
    }

    public function testLanguage(): void
    {
        $object = new BreakElement();
        self::assertNull($object->getLanguage());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->setLanguage('en-US'));
        self::assertNull($object->getLanguage());
    }

    /**
     * Test get/set hash index.
     */
    public function testHashCode(): void
    {
        $object = new BreakElement();
        self::assertEquals(md5(get_class($object)), $object->getHashCode());
    }
}
