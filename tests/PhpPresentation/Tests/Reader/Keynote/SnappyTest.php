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
use PhpOffice\PhpPresentation\Reader\Keynote\Snappy;
use PHPUnit\Framework\TestCase;

/**
 * Test class for the Snappy decompressor of the Keynote reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\Keynote\Snappy
 */
class SnappyTest extends TestCase
{
    public function testDecompressLiteral(): void
    {
        // 5 bytes of output : one literal tag of 5 bytes
        $block = "\x05" . "\x10" . 'Hello';

        self::assertEquals('Hello', Snappy::decompressBlock($block));
    }

    /**
     * A literal longer than 60 bytes states its length in its own bytes.
     */
    public function testDecompressLongLiteral(): void
    {
        $literal = str_repeat('a', 100);
        $block = "\x64" . "\xF0" . "\x63" . $literal;

        self::assertEquals($literal, Snappy::decompressBlock($block));
    }

    /**
     * A copy of one byte, pointing one byte back, is how Snappy writes a run.
     */
    public function testDecompressCopyOverlapping(): void
    {
        // 5 bytes of output : the literal "a", then 4 bytes copied from 1 byte back
        $block = "\x05" . "\x00" . 'a' . "\x01\x01";

        self::assertEquals('aaaaa', Snappy::decompressBlock($block));
    }

    public function testDecompressCopyTwoBytesOffset(): void
    {
        // 8 bytes of output : the literal "abcd", then those 4 bytes again
        $block = "\x08" . "\x0C" . 'abcd' . "\x0E\x04\x00";

        self::assertEquals('abcdabcd', Snappy::decompressBlock($block));
    }

    public function testDecompressCopyFourBytesOffset(): void
    {
        // 8 bytes of output : the literal "abcd", then those 4 bytes again
        $block = "\x08" . "\x0C" . 'abcd' . "\x0F\x04\x00\x00\x00";

        self::assertEquals('abcdabcd', Snappy::decompressBlock($block));
    }

    public function testDecompressBlockWrongLength(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('Snappy block announces 9 bytes but decompresses to 5');

        Snappy::decompressBlock("\x09" . "\x10" . 'Hello');
    }

    public function testDecompressBlockTruncatedLiteral(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('truncated Snappy literal');

        Snappy::decompressBlock("\x05" . "\x10" . 'He');
    }

    public function testDecompressBlockCopyOutOfRange(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('Snappy copy points 8 bytes back into 1 bytes of output');

        Snappy::decompressBlock("\x05" . "\x00" . 'a' . "\x01\x08");
    }

    public function testDecompressBlockMalformedLength(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('malformed Snappy block length');

        Snappy::decompressBlock("\x80\x80\x80\x80\x80");
    }

    /**
     * The frames of a component are decompressed in order and concatenated.
     */
    public function testDecompressFrames(): void
    {
        $first = "\x05" . "\x10" . 'Hello';
        $second = "\x06" . "\x14" . ' world';
        $data = $this->frame($first) . $this->frame($second);

        self::assertEquals('Hello world', Snappy::decompressFrames($data));
    }

    public function testDecompressFramesEmpty(): void
    {
        self::assertEquals('', Snappy::decompressFrames(''));
    }

    public function testDecompressFramesUnsupportedType(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('unsupported Snappy frame type 0x01');

        Snappy::decompressFrames("\x01\x01\x00\x00\x00");
    }

    public function testDecompressFramesTruncatedHeader(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('truncated Snappy frame header');

        Snappy::decompressFrames("\x00\x01\x00");
    }

    public function testDecompressFramesTruncatedFrame(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('truncated Snappy frame');

        Snappy::decompressFrames("\x00\x10\x00\x00" . 'short');
    }

    /**
     * The header Apple writes in front of a block : the type and the length, little endian.
     */
    protected function frame(string $block): string
    {
        return "\x00" . substr(pack('V', strlen($block)), 0, 3) . $block;
    }
}
