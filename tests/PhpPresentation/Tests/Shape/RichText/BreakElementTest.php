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

namespace PhpOffice\PhpPresentation\Tests\Shape\RichText;

use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PHPUnit\Framework\TestCase;

/**
 * Test class for BreakElement element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\RichText\BreakElement
 */
class BreakElementTest extends TestCase
{
    /**
     * Test can read
     */
    public function testText()
    {
        $object = new BreakElement();
        static::assertEquals("\r\n", $object->getText());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->setText());
        static::assertEquals("\r\n", $object->getText());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->setText('AAA'));
        static::assertEquals("\r\n", $object->getText());
    }

    public function testFont()
    {
        $object = new BreakElement();
        static::assertNull($object->getFont());
    }

    public function testLanguage()
    {
        $object = new BreakElement();
        static::assertNull($object->getLanguage());
        static::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\BreakElement', $object->setLanguage('en-US'));
        static::assertNull($object->getLanguage());
    }

    /**
     * Test get/set hash index
     */
    public function testHashCode()
    {
        $object = new BreakElement();
        static::assertEquals(md5(get_class($object)), $object->getHashCode());
    }
}
