<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Common\Compression;

class Snappy
{
    private const CHUNK_TYPE_COMPRESSED = 0x00;
    private const CHUNK_TYPE_UNCOMPRESSED = 0x01;
    private const MAX_BLOCK_SIZE = 65536;
    private const MAX_OFFSET = 32768;

    public function compress(string $data): string
    {
        if (empty($data)) {
            return '';
        }

        $chunks = [];
        $offset = 0;
        $length = strlen($data);

        while ($offset < $length) {
            $blockSize = min(self::MAX_BLOCK_SIZE, $length - $offset);
            $block = substr($data, $offset, $blockSize);

            try {
                $compressed = $this->compressBlock($block);
                $chunks[] = $this->createChunk(
                    $compressed,
                    strlen($compressed) < strlen($block) ? self::CHUNK_TYPE_COMPRESSED : self::CHUNK_TYPE_UNCOMPRESSED
                );
            } catch (\Exception $e) {
                // If compression fails, store uncompressed
                $chunks[] = $this->createChunk($block, self::CHUNK_TYPE_UNCOMPRESSED);
            }

            $offset += $blockSize;
        }

        return implode('', $chunks);
    }

    public function decompress(string $data): string
    {
        if (empty($data)) {
            return '';
        }

        $result = '';
        $offset = 0;
        $length = strlen($data);

        try {
            while ($offset < $length) {
                if ($offset + 5 > $length) {
                    throw new \RuntimeException('Invalid chunk header');
                }

                $header = unpack('Ctype/Vsize', substr($data, $offset, 5));
                if (!$header) {
                    throw new \RuntimeException('Failed to unpack chunk header');
                }

                $offset += 5;
                if ($offset + $header['size'] > $length) {
                    throw new \RuntimeException('Invalid chunk size');
                }

                $chunk = substr($data, $offset, $header['size']);
                $offset += $header['size'];

                $result .= ($header['type'] === self::CHUNK_TYPE_COMPRESSED)
                    ? $this->decompressBlock($chunk)
                    : $chunk;
            }
        } catch (\Exception $e) {
            throw new \RuntimeException('Decompression failed: ' . $e->getMessage());
        }

        return $result;
    }

    private function compressBlock(string $data): string
    {
        if (empty($data)) {
            return '';
        }

        $result = '';
        $length = strlen($data);
        $pos = 0;
        $hashTable = [];

        while ($pos < $length) {
            // Look for matches in the last MAX_OFFSET bytes
            $maxLookback = max(0, $pos - self::MAX_OFFSET);
            $match = $this->findLongestMatch($data, $pos, $maxLookback, $hashTable);

            if ($match && $match['length'] > 3) {
                // Encode match
                $result .= $this->encodeMatch($match['offset'], $match['length']);
                $pos += $match['length'];
            } else {
                // Encode literal
                $literalLength = min(self::MAX_BLOCK_SIZE, $length - $pos);
                $result .= $this->encodeLiteral(substr($data, $pos, $literalLength));
                $pos += $literalLength;
            }

            // Update hash table
            $hashTable[$this->hash(substr($data, $pos, 4))] = $pos;
        }

        return $result;
    }

    private function decompressBlock(string $data): string
    {
        if (empty($data)) {
            return '';
        }

        $result = '';
        $pos = 0;
        $length = strlen($data);

        while ($pos < $length) {
            $tag = ord($data[$pos++]);

            if ($tag & 0x80) {
                // Match
                if ($pos + 1 > $length) {
                    throw new \RuntimeException('Invalid match data');
                }

                $matchLength = (($tag >> 2) & 0x1F) + 4;
                $matchOffset = (ord($data[$pos++]) << 3) | ($tag >> 5);

                if ($matchOffset > strlen($result)) {
                    throw new \RuntimeException('Invalid match offset');
                }

                // Copy from back reference
                $start = strlen($result) - $matchOffset;
                for ($i = 0; $i < $matchLength; $i++) {
                    $result .= $result[$start + $i];
                }
            } else {
                // Literal
                $literalLength = ($tag & 0x7F) + 1;
                if ($pos + $literalLength > $length) {
                    throw new \RuntimeException('Invalid literal length');
                }

                $result .= substr($data, $pos, $literalLength);
                $pos += $literalLength;
            }
        }

        return $result;
    }

    private function createChunk(string $data, int $type): string
    {
        return pack('CV', $type, strlen($data)) . $data;
    }

    private function hash(string $data): int
    {
        // Simple rolling hash function
        $hash = 0;
        for ($i = 0; $i < min(4, strlen($data)); $i++) {
            $hash = ($hash * 33) + ord($data[$i]);
        }
        return $hash & 0xFFFFFFFF;
    }

    private function findLongestMatch(string $data, int $pos, int $maxLookback, array $hashTable): ?array
    {
        $length = strlen($data);
        if ($pos + 4 > $length) {
            return null;
        }

        $hash = $this->hash(substr($data, $pos, 4));
        if (!isset($hashTable[$hash]) || $hashTable[$hash] < $maxLookback) {
            return null;
        }

        $matchPos = $hashTable[$hash];
        $matchLength = 0;
        while (
            $pos + $matchLength < $length &&
            $matchLength < 255 &&
            $data[$matchPos + $matchLength] === $data[$pos + $matchLength]
        ) {
            $matchLength++;
        }

        return $matchLength >= 4 ? [
            'offset' => $pos - $matchPos,
            'length' => $matchLength
        ] : null;
    }

    private function encodeLiteral(string $literal): string
    {
        $length = strlen($literal) - 1;
        return chr($length) . $literal;
    }

    private function encodeMatch(int $offset, int $length): string
    {
        $tag = 0x80 | (($length - 4) << 2) | ($offset >> 3);
        return chr($tag) . chr($offset & 0xFF);
    }
}
