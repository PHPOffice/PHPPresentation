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

use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Style\Font;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Run element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\RichText\Run
 */
class RunTest extends TestCase
{
    /**
     * Test can read.
     */
    public function testConstruct(): void
    {
        $object = new Run();
        self::assertEquals('', $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());

        $object = new Run('BBB');
        self::assertEquals('BBB', $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
    }

    public function testFont(): void
    {
        $object = new Run();
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->setFont(new Font()));
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
    }

    public function testLanguage(): void
    {
        $object = new Run();
        self::assertNull($object->getLanguage());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->setLanguage('en-US'));
        self::assertEquals('en-US', $object->getLanguage());
    }

    public function testText(): void
    {
        $object = new Run();
        self::assertEquals('', $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->setText());
        self::assertEquals('', $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\RichText\\Run', $object->setText('AAA'));
        self::assertEquals('AAA', $object->getText());

        $object = new Run('BBB');
        self::assertEquals('BBB', $object->getText());
    }

    /**
     * Test get/set hash index.
     */
    public function testHashCode(): void
    {
        $object = new Run();
        self::assertEquals(md5($object->getFont()->getHashCode() . get_class($object)), $object->getHashCode());
    }
}
