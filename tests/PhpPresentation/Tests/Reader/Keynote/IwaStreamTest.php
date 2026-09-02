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
use PhpOffice\PhpPresentation\Reader\Keynote\IwaStream;
use PHPUnit\Framework\TestCase;

/**
 * Test class for the IWA stream reader of the Keynote reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\Keynote\IwaStream
 */
class IwaStreamTest extends TestCase
{
    /**
     * Every archive of a stream is read, with the messages its header announces.
     */
    public function testParse(): void
    {
        $data = $this->archive(1, [[2001, 'the text']])
            . $this->archive(2, [[6, 'a'], [11, 'bb']]);

        self::assertEquals([
            [
                'identifier' => 1,
                'messages' => [
                    ['type' => 2001, 'payload' => 'the text'],
                ],
            ],
            [
                'identifier' => 2,
                'messages' => [
                    ['type' => 6, 'payload' => 'a'],
                    ['type' => 11, 'payload' => 'bb'],
                ],
            ],
        ], IwaStream::parse($data));
    }

    public function testParseEmpty(): void
    {
        self::assertEquals([], IwaStream::parse(''));
    }

    /**
     * An archive announcing no message is read, and holds none.
     */
    public function testParseArchiveWithoutMessage(): void
    {
        self::assertEquals(
            [['identifier' => 7, 'messages' => []]],
            IwaStream::parse($this->archive(7, []))
        );
    }

    public function testParseTruncatedHeader(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('truncated IWA archive header');

        IwaStream::parse("\x20\x08\x01");
    }

    public function testParseTruncatedMessage(): void
    {
        $data = $this->archive(1, [[2001, 'the text']]);

        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage('truncated IWA message');

        IwaStream::parse(substr($data, 0, -4));
    }

    /**
     * An archive of a stream : the length of its `TSP.ArchiveInfo`, that message, then the payloads.
     *
     * @param array<int, array{0: int, 1: string}> $messages
     */
    protected function archive(int $identifier, array $messages): string
    {
        $info = "\x08" . $this->varint($identifier);
        $payloads = '';

        foreach ($messages as [$type, $payload]) {
            $messageInfo = "\x08" . $this->varint($type)
                . "\x10" . $this->varint(1)
                . "\x18" . $this->varint(strlen($payload));
            $info .= "\x12" . $this->varint(strlen($messageInfo)) . $messageInfo;
            $payloads .= $payload;
        }

        return $this->varint(strlen($info)) . $info . $payloads;
    }

    /**
     * A base 128 integer.
     */
    protected function varint(int $value): string
    {
        $bytes = '';
        do {
            $byte = $value & 0x7F;
            $value >>= 7;
            $bytes .= chr($byte | ($value > 0 ? 0x80 : 0x00));
        } while ($value > 0);

        return $bytes;
    }
}
