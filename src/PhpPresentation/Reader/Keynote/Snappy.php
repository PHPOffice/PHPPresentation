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
 * The Snappy decompressor an IWA stream asks for.
 *
 * An `.iwa` component is a series of frames : a one byte type, a three byte little endian length
 * and that many bytes of a Snappy compressed block. Apple writes only the uncompressed type and
 * neither the stream identifier nor the checksums the official framing format spells out, so the
 * frames are read here rather than by a Snappy library.
 */
class Snappy
{
    /**
     * The only frame type Apple writes.
     *
     * @var int
     */
    protected const FRAME_TYPE_COMPRESSED = 0x00;

    /**
     * Every frame of a component, decompressed and concatenated.
     *
     * @param string $data Content of an `.iwa` component
     * @param string $context Name of the component, for error messages
     */
    public static function decompressFrames(string $data, string $context = ''): string
    {
        $output = '';
        $offset = 0;
        $length = strlen($data);

        while ($offset < $length) {
            if ($offset + 4 > $length) {
                throw new InvalidFileFormatException($context, self::class, 'truncated Snappy frame header');
            }
            $type = ord($data[$offset]);
            if (self::FRAME_TYPE_COMPRESSED !== $type) {
                throw new InvalidFileFormatException($context, self::class, sprintf('unsupported Snappy frame type 0x%02X', $type));
            }
            $frameLength = ord($data[$offset + 1]) | ord($data[$offset + 2]) << 8 | ord($data[$offset + 3]) << 16;
            $offset += 4;
            if ($offset + $frameLength > $length) {
                throw new InvalidFileFormatException($context, self::class, 'truncated Snappy frame');
            }
            $output .= self::decompressBlock(substr($data, $offset, $frameLength), $context);
            $offset += $frameLength;
        }

        return $output;
    }

    /**
     * A single Snappy compressed block, as described by the format specification.
     *
     * @param string $block A length prefixed Snappy block
     * @param string $context Name of the component, for error messages
     */
    public static function decompressBlock(string $block, string $context = ''): string
    {
        $offset = 0;
        $expectedLength = self::readVarint($block, $offset, $context);
        $length = strlen($block);
        $output = '';

        while ($offset < $length) {
            $tag = ord($block[$offset]);
            ++$offset;
            if (0 === ($tag & 0x03)) {
                $output .= self::readLiteral($block, $offset, $tag, $context);
            } else {
                self::readCopy($block, $offset, $tag, $output, $context);
            }
        }

        if (strlen($output) !== $expectedLength) {
            throw new InvalidFileFormatException($context, self::class, sprintf(
                'Snappy block announces %d bytes but decompresses to %d',
                $expectedLength,
                strlen($output)
            ));
        }

        return $output;
    }

    /**
     * The bytes a literal tag copies straight out of the block.
     */
    protected static function readLiteral(string $block, int &$offset, int $tag, string $context): string
    {
        $literalLength = $tag >> 2;
        if ($literalLength >= 60) {
            $sizeBytes = $literalLength - 59;
            $literalLength = self::readLittleEndian($block, $offset, $sizeBytes, $context);
            $offset += $sizeBytes;
        }
        ++$literalLength;

        if ($offset + $literalLength > strlen($block)) {
            throw new InvalidFileFormatException($context, self::class, 'truncated Snappy literal');
        }
        $literal = substr($block, $offset, $literalLength);
        $offset += $literalLength;

        return $literal;
    }

    /**
     * What a copy tag repeats from the bytes already decompressed. The copy may overlap the end of
     * the output, which is what makes a run of one byte cheap to write, so it is byte by byte.
     */
    protected static function readCopy(string $block, int &$offset, int $tag, string &$output, string $context): void
    {
        switch ($tag & 0x03) {
            case 0x01:
                $copyLength = 4 + (($tag >> 2) & 0x07);
                $copyOffset = (($tag >> 5) & 0x07) << 8 | self::readLittleEndian($block, $offset, 1, $context);
                ++$offset;

                break;
            case 0x02:
                $copyLength = ($tag >> 2) + 1;
                $copyOffset = self::readLittleEndian($block, $offset, 2, $context);
                $offset += 2;

                break;
            default:
                $copyLength = ($tag >> 2) + 1;
                $copyOffset = self::readLittleEndian($block, $offset, 4, $context);
                $offset += 4;

                break;
        }

        $written = strlen($output);
        if ($copyOffset < 1 || $copyOffset > $written) {
            throw new InvalidFileFormatException($context, self::class, sprintf(
                'Snappy copy points %d bytes back into %d bytes of output',
                $copyOffset,
                $written
            ));
        }

        $start = $written - $copyOffset;
        for ($index = 0; $index < $copyLength; ++$index) {
            $output .= $output[$start + $index];
        }
    }

    /**
     * An unsigned little endian integer of the given width.
     */
    protected static function readLittleEndian(string $block, int $offset, int $sizeBytes, string $context): int
    {
        if ($offset + $sizeBytes > strlen($block)) {
            throw new InvalidFileFormatException($context, self::class, 'truncated Snappy tag');
        }

        $value = 0;
        for ($index = 0; $index < $sizeBytes; ++$index) {
            $value |= ord($block[$offset + $index]) << (8 * $index);
        }

        return $value;
    }

    /**
     * The base 128 length a Snappy block opens with.
     */
    protected static function readVarint(string $block, int &$offset, string $context): int
    {
        $value = 0;
        $shift = 0;
        $length = strlen($block);

        while ($offset < $length) {
            $byte = ord($block[$offset]);
            ++$offset;
            $value |= ($byte & 0x7F) << $shift;
            if (0 === ($byte & 0x80)) {
                return $value;
            }
            $shift += 7;
            if ($shift > 28) {
                break;
            }
        }

        throw new InvalidFileFormatException($context, self::class, 'malformed Snappy block length');
    }
}
