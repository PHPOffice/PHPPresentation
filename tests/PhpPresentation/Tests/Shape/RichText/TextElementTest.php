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

use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PHPUnit\Framework\TestCase;

/**
 * Test class for TextElement element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\RichText\TextElement
 */
class TextElementTest extends TestCase
{
    /**
     * Test can read.
     */
    public function testConstruct(): void
    {
        $object = new TextElement();
        self::assertEquals('', $object->getText());

        $object = new TextElement('AAA');
        self::assertEquals('AAA', $object->getText());
    }

    public function testFont(): void
    {
        $object = new TextElement();
        self::assertNull($object->getFont());
    }

    public function testHyperlink(): void
    {
        $object = new TextElement();
        self::assertFalse($object->hasHyperlink());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setHyperlink());
        self::assertFalse($object->hasHyperlink());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->getHyperlink());
        self::assertTrue($object->hasHyperlink());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setHyperlink(new Hyperlink('http://www.google.fr')));
        self::assertTrue($object->hasHyperlink());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->getHyperlink());
    }

    public function testLanguage(): void
    {
        $object = new TextElement();
        self::assertNull($object->getLanguage());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setLanguage('en-US'));
        self::assertEquals('en-US', $object->getLanguage());
    }

    public function testText(): void
    {
        $object = new TextElement();
        self::assertEquals('', $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setText());
        self::assertEquals('', $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\TextElement', $object->setText('AAA'));
        self::assertEquals('AAA', $object->getText());
    }

    /**
     * Test get/set hash index.
     */
    public function testHashCode(): void
    {
        $object = new TextElement();
        self::assertEquals(md5(get_class($object)), $object->getHashCode());
    }
}
