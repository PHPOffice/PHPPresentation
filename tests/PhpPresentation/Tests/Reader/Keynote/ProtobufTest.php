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

namespace PhpOffice\PhpPresentation\Tests\Reader\Keynote;

use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\Reader\Keynote\Protobuf;
use PHPUnit\Framework\TestCase;

/**
 * Test class for the Protobuf wire format reader of the Keynote reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\Keynote\Protobuf
 */
class ProtobufTest extends TestCase
{
    public function testDecodeVarint(): void
    {
        // field 1, varint 300
        $fields = Protobuf::decode("\x08\xAC\x02");

        self::assertEquals([1 => [300]], $fields);
        self::assertEquals(300, Protobuf::getInt($fields, 1));
        self::assertNull(Protobuf::getInt($fields, 2));
    }

    public function testDecodeLengthDelimited(): void
    {
        // field 3, the bytes "Keynote"
        $fields = Protobuf::decode("\x1A\x07Keynote");

        self::assertEquals([3 => ['Keynote']], $fields);
        self::assertEquals(['Keynote'], Protobuf::getStrings($fields, 3));
        self::assertEquals([], Protobuf::getStrings($fields, 1));
        self::assertNull(Protobuf::getInt($fields, 3));
    }

    /**
     * A field said twice keeps both of its values, in the order the message states them.
     */
    public function testDecodeRepeatedField(): void
    {
        $fields = Protobuf::decode("\x1A\x03one\x1A\x03two");

        self::assertEquals(['one', 'two'], Protobuf::getStrings($fields, 3));
    }

    public function testDecodeFixedWidth(): void
    {
        // field 1 as a fixed 64, field 2 as a fixed 32
        $fields = Protobuf::decode("\x09" . '01234567' . "\x15" . '89AB');

        self::assertEquals('01234567', $fields[1][0]);
        self::assertEquals('89AB', $fields[2][0]);
    }

    public function testDecodeEmpty(): void
    {
        self::assertEquals([], Protobuf::decode(''));
    }

    public function testDecodeUnsupportedWireType(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('unsupported Protobuf wire type 3');

        Protobuf::decode("\x0B");
    }

    public function testDecodeFieldNumberZero(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('Protobuf field number 0');

        Protobuf::decode("\x00\x01");
    }

    public function testDecodeTruncatedField(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('truncated Protobuf field');

        Protobuf::decode("\x1A\x07Key");
    }

    public function testDecodeMalformedVarint(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('malformed Protobuf varint');

        Protobuf::decode("\x08\x80\x80");
    }

    public function testTryDecode(): void
    {
        self::assertEquals([3 => ['Keynote']], Protobuf::tryDecode("\x1A\x07Keynote"));
        self::assertNull(Protobuf::tryDecode("\x1A\x07Key"));
    }
}
