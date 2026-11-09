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

namespace PhpOffice\PhpPresentation\Reader\Serialized;

use ArrayIterator;
use ArrayObject;
use UnexpectedValueException;

/**
 * Validate the serialized grammar before PHP allocates objects or runs handlers.
 */
final class DataValidator
{
    /** @var string */
    private $data;

    /** @var int */
    private $offset = 0;

    /** @var array<string, bool> */
    private $classes;

    /** @var int */
    private $remaining;

    /** @var int */
    private $maxDepth;

    /** @param array<class-string> $classes */
    public function __construct(string $data, array $classes, int $maxValues, int $maxDepth)
    {
        $this->data = $data;
        $this->classes = array_fill_keys($classes, true);
        $this->remaining = $maxValues;
        $this->maxDepth = $maxDepth;
    }

    public function validate(): void
    {
        $this->value(0);
        if ($this->offset !== strlen($this->data)) {
            throw new UnexpectedValueException('Trailing serialized data');
        }
    }

    /**
     * Return string/integer values only, for checking keys and iterator class names.
     */
    private function value(int $depth, bool $key = false): ?string
    {
        if (--$this->remaining < 0 || $depth > $this->maxDepth) {
            throw new UnexpectedValueException('Serialized data exceeds resource limits');
        }
        $type = $this->data[$this->offset] ?? '';
        ++$this->offset;
        if ($key && $type !== 'i' && $type !== 's') {
            throw new UnexpectedValueException('Invalid serialized key');
        }
        if ($type === 'N') {
            $this->expect(';');

            return null;
        }
        $this->expect(':');
        if ($type === 's') {
            return $this->string(';');
        }
        if (in_array($type, ['i', 'b', 'd', 'r', 'R'], true)) {
            $number = $this->until(';');
            $pattern = $type === 'd' ? '/^(?:NAN|-?INF|[+-]?[0-9]+(?:\.[0-9]*)?(?:[Ee][+-]?[0-9]+)?)$/D' : '/^-?[0-9]+$/D';
            if (!preg_match($pattern, $number) || ($type === 'b' && $number !== '0' && $number !== '1')) {
                throw new UnexpectedValueException('Invalid serialized scalar');
            }

            return $type === 'i' ? $number : null;
        }
        if ($type !== 'a' && $type !== 'O') {
            // Reject opaque custom payloads and enums before handlers or autoloading run.
            throw new UnexpectedValueException('Unsupported serialized value');
        }
        $class = null;
        if ($type === 'O') {
            $class = $this->string(':');
            if (!isset($this->classes[$class])) {
                throw new UnexpectedValueException('Forbidden serialized class');
            }
        }
        $count = $this->count();
        $this->expect('{');
        if ($count > intdiv($this->remaining, 2)) {
            throw new UnexpectedValueException('Serialized collection exceeds resource limits');
        }
        if ($class === ArrayObject::class && $count !== 4) {
            throw new UnexpectedValueException('Invalid ArrayObject state');
        }
        for ($i = 0; $i < $count; ++$i) {
            $name = $this->value($depth + 1, true);
            $valueType = $this->data[$this->offset] ?? '';
            $value = $this->value($depth + 1);
            if ($class === ArrayObject::class) {
                // Its internal handler resolves the iterator class during unserialization.
                if ($name !== (string) $i
                    || ($i === 0 && $valueType !== 'i')
                    || (($i === 1 || $i === 2) && $valueType !== 'a')
                    || ($i === 3 && $valueType !== 'N' && !($valueType === 's' && $value === ArrayIterator::class))) {
                    throw new UnexpectedValueException('Unsafe ArrayObject state');
                }
            }
        }
        $this->expect('}');

        return null;
    }

    private function string(string $terminator): string
    {
        $length = $this->count();
        $this->expect('"');
        if ($length > strlen($this->data) - $this->offset) {
            throw new UnexpectedValueException('Invalid serialized string length');
        }
        $value = substr($this->data, $this->offset, $length);
        $this->offset += $length;
        $this->expect('"' . $terminator);

        return $value;
    }

    private function count(): int
    {
        $number = $this->until(':');
        if (!preg_match('/^[0-9]+$/D', $number) || strlen($number) > 9) {
            throw new UnexpectedValueException('Invalid serialized count');
        }

        return (int) $number;
    }

    private function until(string $delimiter): string
    {
        $end = strpos($this->data, $delimiter, $this->offset);
        if ($end === false) {
            throw new UnexpectedValueException('Truncated serialized data');
        }
        $value = substr($this->data, $this->offset, $end - $this->offset);
        $this->offset = $end + 1;

        return $value;
    }

    private function expect(string $text): void
    {
        if (substr($this->data, $this->offset, strlen($text)) !== $text) {
            throw new UnexpectedValueException('Malformed serialized data');
        }
        $this->offset += strlen($text);
    }
}
