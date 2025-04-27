<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Common\Adapter\Protobuf;

class ProtobufAdapter
{
    public function encode(array $data): string
    {
        // Implement Protobuf encoding
        // This is a placeholder
        return serialize($data);
    }

    public function decode(string $data): array
    {
        // Implement Protobuf decoding
        // This is a placeholder
        return unserialize($data);
    }

    protected function encodeVarint(int $value): string
    {
        $bytes = '';
        while ($value > 0x7F) {
            $bytes .= chr(($value & 0x7F) | 0x80);
            $value >>= 7;
        }
        $bytes .= chr($value & 0x7F);
        return $bytes;
    }

    protected function decodeVarint(string $data, int &$offset): int
    {
        $value = 0;
        $shift = 0;
        while (true) {
            $byte = ord($data[$offset++]);
            $value |= ($byte & 0x7F) << $shift;
            if (($byte & 0x80) === 0) {
                break;
            }
            $shift += 7;
        }
        return $value;
    }
}
