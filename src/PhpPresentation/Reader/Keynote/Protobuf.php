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

namespace PhpOffice\PhpPresentation\Reader\Keynote;

use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;

/**
 * As much of the Protocol Buffers wire format as an IWA stream needs.
 *
 * A message is read without its schema : every field is kept under its number, a varint as an
 * integer and a length delimited field as the bytes it carries. That is enough to reach the text
 * and the file names Keynote stores, and it keeps the reader free of a Protobuf dependency and of
 * the hundreds of Apple messages a generated schema would bring.
 */
class Protobuf
{
    /**
     * @var int
     */
    public const WIRE_VARINT = 0;

    /**
     * @var int
     */
    public const WIRE_FIXED64 = 1;

    /**
     * @var int
     */
    public const WIRE_LENGTH_DELIMITED = 2;

    /**
     * @var int
     */
    public const WIRE_FIXED32 = 5;

    /**
     * Every field of a message, by field number. A varint is an integer, anything else the bytes
     * it carries.
     *
     * @param string $data Bytes of a single message
     * @param string $context Name of the component, for error messages
     *
     * @return array<int, array<int, int|string>>
     */
    public static function decode(string $data, string $context = ''): array
    {
        $fields = [];
        $offset = 0;
        $length = strlen($data);

        while ($offset < $length) {
            $key = self::readVarint($data, $offset, $context);
            $number = $key >> 3;
            if ($number < 1) {
                throw new InvalidFileFormatException($context, self::class, 'Protobuf field number 0');
            }
            $fields[$number][] = self::readValue($data, $offset, $key & 0x07, $context);
        }

        return $fields;
    }

    /**
     * The same as decode(), for bytes which may not be a message at all.
     *
     * @return null|array<int, array<int, int|string>>
     */
    public static function tryDecode(string $data): ?array
    {
        try {
            return self::decode($data);
        } catch (InvalidFileFormatException $e) {
            return null;
        }
    }

    /**
     * Every length delimited value of a field, as bytes.
     *
     * @param array<int, array<int, int|string>> $fields
     *
     * @return array<int, string>
     */
    public static function getStrings(array $fields, int $number): array
    {
        return array_values(array_filter($fields[$number] ?? [], 'is_string'));
    }

    /**
     * The first varint value of a field.
     *
     * @param array<int, array<int, int|string>> $fields
     */
    public static function getInt(array $fields, int $number): ?int
    {
        foreach ($fields[$number] ?? [] as $value) {
            if (is_int($value)) {
                return $value;
            }
        }

        return null;
    }

    /**
     * The value a field of the given wire type carries.
     *
     * @return int|string
     */
    protected static function readValue(string $data, int &$offset, int $wireType, string $context)
    {
        switch ($wireType) {
            case self::WIRE_VARINT:
                return self::readVarint($data, $offset, $context);
            case self::WIRE_FIXED64:
                return self::readBytes($data, $offset, 8, $context);
            case self::WIRE_FIXED32:
                return self::readBytes($data, $offset, 4, $context);
            case self::WIRE_LENGTH_DELIMITED:
                return self::readBytes($data, $offset, self::readVarint($data, $offset, $context), $context);
            default:
                throw new InvalidFileFormatException($context, self::class, sprintf('unsupported Protobuf wire type %d', $wireType));
        }
    }

    /**
     * The next bytes of the message.
     */
    protected static function readBytes(string $data, int &$offset, int $length, string $context): string
    {
        if ($length < 0 || $offset + $length > strlen($data)) {
            throw new InvalidFileFormatException($context, self::class, 'truncated Protobuf field');
        }
        $bytes = substr($data, $offset, $length);
        $offset += $length;

        return $bytes;
    }

    /**
     * A base 128 integer.
     */
    public static function readVarint(string $data, int &$offset, string $context = ''): int
    {
        $value = 0;
        $shift = 0;
        $length = strlen($data);

        while ($offset < $length) {
            $byte = ord($data[$offset]);
            ++$offset;
            $value |= ($byte & 0x7F) << $shift;
            if (0 === ($byte & 0x80)) {
                return $value;
            }
            $shift += 7;
            if ($shift > 63) {
                break;
            }
        }

        throw new InvalidFileFormatException($context, self::class, 'malformed Protobuf varint');
    }
}
