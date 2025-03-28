<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Common\Protobuf;

class Message
{
    private const WIRE_VARINT = 0;
    private const WIRE_64BIT = 1;
    private const WIRE_LENGTH_DELIMITED = 2;
    private const WIRE_32BIT = 5;

    public function encode(array $data): string
    {
        try {
            $output = '';
            foreach ($data as $field => $value) {
                if (!is_int($field)) {
                    throw new \InvalidArgumentException('Field number must be integer');
                }

                $wireType = $this->getWireType($value);
                $output .= $this->encodeTag($field, $wireType);
                $output .= $this->encodeValue($value, $wireType);
            }
            return $output;
        } catch (\Exception $e) {
            throw new \RuntimeException('Protobuf encoding failed: ' . $e->getMessage());
        }
    }

    public function decode(string $data): array
    {
        try {
            $result = [];
            $offset = 0;
            $length = strlen($data);

            while ($offset < $length) {
                $tag = $this->decodeVarint($data, $offset);
                if ($tag === false) {
                    throw new \RuntimeException('Invalid varint encoding');
                }

                $wireType = $tag & 0x07;
                $fieldNumber = $tag >> 3;

                $value = $this->decodeValue($data, $offset, $wireType);
                if ($value !== null) {
                    $result[$fieldNumber] = $value;
                }
            }

            return $result;
        } catch (\Exception $e) {
            throw new \RuntimeException('Protobuf decoding failed: ' . $e->getMessage());
        }
    }

    private function getWireType($value): int
    {
        if (is_int($value)) {
            return self::WIRE_VARINT;
        }
        if (is_string($value)) {
            return self::WIRE_LENGTH_DELIMITED;
        }
        if (is_float($value)) {
            return self::WIRE_64BIT;
        }
        if (is_array($value)) {
            return self::WIRE_LENGTH_DELIMITED;
        }
        throw new \InvalidArgumentException('Unsupported value type');
    }

    private function encodeTag(int $fieldNumber, int $wireType): string
    {
        if ($fieldNumber <= 0) {
            throw new \InvalidArgumentException('Field number must be positive');
        }
        return $this->encodeVarint(($fieldNumber << 3) | $wireType);
    }

    private function encodeValue($value, int $wireType): string
    {
        switch ($wireType) {
            case self::WIRE_VARINT:
                return $this->encodeVarint($value);
            case self::WIRE_64BIT:
                return $this->encode64Bit($value);
            case self::WIRE_LENGTH_DELIMITED:
                if (is_array($value)) {
                    // Recursively encode nested messages
                    $encodedValue = $this->encode($value);
                    return $this->encodeLengthDelimited($encodedValue);
                }
                return $this->encodeLengthDelimited($value);
            case self::WIRE_32BIT:
                return $this->encode32Bit($value);
            default:
                throw new \InvalidArgumentException('Unknown wire type');
        }
    }

    private function decodeValue(string $data, int &$offset, int $wireType)
    {
        switch ($wireType) {
            case self::WIRE_VARINT:
                return $this->decodeVarint($data, $offset);
            case self::WIRE_64BIT:
                return $this->decode64Bit($data, $offset);
            case self::WIRE_LENGTH_DELIMITED:
                return $this->decodeLengthDelimited($data, $offset);
            case self::WIRE_32BIT:
                return $this->decode32Bit($data, $offset);
            default:
                // Skip unknown wire types
                if ($wireType === 3) { // WIRE_START_GROUP
                    return null;
                }
                if ($wireType === 4) { // WIRE_END_GROUP
                    return null;
                }
                // For other unknown types, try to decode as varint
                return $this->decodeVarint($data, $offset);
        }
    }

    private function encodeVarint(int $value): string
    {
        $output = '';
        while ($value > 0x7F) {
            $output .= chr(($value & 0x7F) | 0x80);
            $value >>= 7;
        }
        $output .= chr($value & 0x7F);
        return $output;
    }

    private function decodeVarint(string $data, int &$offset): ?int
    {
        $value = 0;
        $shift = 0;

        while ($offset < strlen($data)) {
            $byte = ord($data[$offset++]);
            $value |= ($byte & 0x7F) << $shift;
            if (($byte & 0x80) === 0) {
                return $value;
            }
            $shift += 7;
            if ($shift >= 64) {
                throw new \RuntimeException('Varint is too long');
            }
        }

        return null;
    }

    private function encodeLengthDelimited(string $value): string
    {
        return $this->encodeVarint(strlen($value)) . $value;
    }

    private function decodeLengthDelimited(string $data, int &$offset): mixed
    {
        $length = $this->decodeVarint($data, $offset);
        if ($length === null) {
            return '';
        }
        if ($offset + $length > strlen($data)) {
            $length = strlen($data) - $offset;
        }
        if ($length <= 0) {
            return '';
        }

        $value = substr($data, $offset, $length);
        $offset += $length;

        // Try to decode as a nested message
        try {
            return $this->decode($value);
        } catch (\Exception $e) {
            // If decoding as a nested message fails, return as string
            return $value;
        }
    }

    private function encode64Bit(float $value): string
    {
        return pack('d', $value);
    }

    private function decode64Bit(string $data, int &$offset): float
    {
        if ($offset + 8 > strlen($data)) {
            throw new \RuntimeException('Invalid 64-bit value');
        }

        $value = unpack('d', substr($data, $offset, 8))[1];
        $offset += 8;
        return $value;
    }

    private function encode32Bit(float $value): string
    {
        return pack('f', $value);
    }

    private function decode32Bit(string $data, int &$offset): float
    {
        if ($offset + 4 > strlen($data)) {
            throw new \RuntimeException('Invalid 32-bit value');
        }

        $value = unpack('f', substr($data, $offset, 4))[1];
        $offset += 4;
        return $value;
    }
}
