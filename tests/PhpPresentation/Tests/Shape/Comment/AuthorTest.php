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

namespace PhpOffice\PhpPresentation\Tests\Shape\Comment;

use PhpOffice\PhpPresentation\Shape\Comment\Author;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Author element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Comment\Author
 */
class AuthorTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Author();

        self::assertNull($object->getName());
        self::assertNull($object->getIndex());
        self::assertNull($object->getInitials());
    }

    public function testGetSetIndex(): void
    {
        $expectedVal = mt_rand(1, 100);

        $object = new Author();
        self::assertNull($object->getIndex());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment\\Author', $object->setIndex($expectedVal));
        self::assertEquals($expectedVal, $object->getIndex());
    }

    public function testGetSetInitials(): void
    {
        $expectedVal = 'AABBCCDD';

        $object = new Author();
        self::assertNull($object->getInitials());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment\\Author', $object->setInitials($expectedVal));
        self::assertEquals($expectedVal, $object->getInitials());
    }

    public function testGetSetName(): void
    {
        $expectedVal = 'AABBCCDD';

        $object = new Author();
        self::assertNull($object->getName());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Comment\\Author', $object->setName($expectedVal));
        self::assertEquals($expectedVal, $object->getName());
    }
}
