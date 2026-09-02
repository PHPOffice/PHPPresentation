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

use PhpOffice\PhpPresentation\Shape\RichText\Field;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Field element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\RichText\Field
 */
class FieldTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Field();
        self::assertEquals(Field::TYPE_SLIDENUM, $object->getType());
        self::assertEquals('', $object->getText());
        self::assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());

        $object = new Field(Field::TYPE_SLIDECOUNT, '12');
        self::assertEquals(Field::TYPE_SLIDECOUNT, $object->getType());
        self::assertEquals('12', $object->getText());
    }

    public function testType(): void
    {
        $object = new Field();
        // the type is an open string: a name no application knows is still a legal field
        self::assertInstanceOf(Field::class, $object->setType('datetime13'));
        self::assertEquals('datetime13', $object->getType());
    }

    public function testHashCode(): void
    {
        $object = new Field(Field::TYPE_SLIDENUM, 'AAA');
        // two fields saying the same thing about different things are not the same field
        self::assertNotEquals(
            $object->getHashCode(),
            (new Field(Field::TYPE_SLIDECOUNT, 'AAA'))->getHashCode()
        );
        self::assertEquals(
            $object->getHashCode(),
            (new Field(Field::TYPE_SLIDENUM, 'AAA'))->getHashCode()
        );
    }

    public function testParagraphCreateField(): void
    {
        $object = new Paragraph();
        $object->getFont()->setName('Georgia');

        $field = $object->createField(Field::TYPE_DATETIME, '03-04-05');
        self::assertEquals(Field::TYPE_DATETIME, $field->getType());
        self::assertEquals('03-04-05', $field->getText());
        // a field takes the paragraph's font, the same way a run does
        self::assertEquals('Georgia', $field->getFont()->getName());
        self::assertEquals([$field], $object->getRichTextElements());
    }
}
