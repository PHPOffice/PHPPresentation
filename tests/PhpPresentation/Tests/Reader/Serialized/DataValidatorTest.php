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

namespace PhpOffice\PhpPresentation\Tests\Reader\Serialized;

use ArrayObject;
use PhpOffice\PhpPresentation\Reader\Serialized\DataValidator;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use UnexpectedValueException;

class DataValidatorTest extends TestCase
{
    /** @return array<string, array{string}> */
    public static function validValues(): array
    {
        $shared = new ArrayObject(['binary' => "\0\xff", 'unicode' => '文本']);
        $values = [null, true, false, 0, -42, 1.5, -0.0, 1.0e100, INF, -INF, NAN,
            '', "\0\xff", '文本', [], ['nested' => [1, 2]], [$shared, $shared]];
        $cases = [];
        foreach ($values as $i => $value) {
            $cases['value ' . $i] = [serialize($value)];
        }
        $cases['explicit standard iterator'] = ['O:11:"ArrayObject":4:{i:0;i:0;i:1;a:0:{}i:2;a:0:{}i:3;s:13:"ArrayIterator";}'];
        $cases['numeric string state key'] = ['O:11:"ArrayObject":4:{s:1:"0";i:0;i:1;a:0:{}i:2;a:0:{}i:3;N;}'];
        $cases['shared array reference'] = ['a:2:{i:0;a:0:{}i:1;R:2;}'];

        return $cases;
    }

    /** @dataProvider validValues */
    #[DataProvider('validValues')]
    public function testAcceptsValidGrammar(string $data): void
    {
        $this->expectNotToPerformAssertions();
        (new DataValidator($data, [ArrayObject::class], 100, 10))->validate();
    }

    /** @return array<string, array{string, string}> */
    public static function malformedValues(): array
    {
        return [
            'empty input' => ['', 'Malformed serialized data'],
            'missing null terminator' => ['N', 'Malformed serialized data'],
            'invalid null terminator' => ['N:', 'Malformed serialized data'],
            'missing scalar terminator' => ['i:1', 'Truncated serialized data'],
            'boolean out of range' => ['b:2;', 'Invalid serialized scalar'],
            'negative boolean' => ['b:-1;', 'Invalid serialized scalar'],
            'integer with fraction' => ['i:1.5;', 'Invalid serialized scalar'],
            'non-numeric double' => ['d:abc;', 'Invalid serialized scalar'],
            'invalid object reference' => ['r:abc;', 'Invalid serialized scalar'],
            'invalid array reference' => ['R:abc;', 'Invalid serialized scalar'],
            'custom serialization' => ['C:11:"ArrayObject":0:{}', 'Unsupported serialized value'],
            'enum' => ['E:4:"X:Y";', 'Unsupported serialized value'],
            'unknown token' => ['x:1;', 'Unsupported serialized value'],
            'forbidden class' => ['O:8:"stdClass":0:{}', 'Forbidden serialized class'],
            'negative string count' => ['s:-1:"";', 'Invalid serialized count'],
            'non-numeric count' => ['a:x:{}', 'Invalid serialized count'],
            'oversized count syntax' => ['a:1000000000:{}', 'Invalid serialized count'],
            'truncated count' => ['a:1', 'Truncated serialized data'],
            'truncated string' => ['s:10:"hi";', 'Invalid serialized string length'],
            'wrong string length' => ['s:1:"hi";', 'Malformed serialized data'],
            'missing string quotation' => ['s:2:hi;', 'Malformed serialized data'],
            'missing collection opening' => ['a:0:}', 'Malformed serialized data'],
            'missing collection closing' => ['a:0:{', 'Malformed serialized data'],
            'boolean key' => ['a:1:{b:0;N;}', 'Invalid serialized key'],
            'null key' => ['a:1:{N;N;}', 'Invalid serialized key'],
            'collection key' => ['a:1:{a:0:{}N;}', 'Invalid serialized key'],
            'trailing value' => ['N;N;', 'Trailing serialized data'],
            'short ArrayObject state' => ['O:11:"ArrayObject":0:{}', 'Invalid ArrayObject state'],
            'long ArrayObject state' => ['O:11:"ArrayObject":5:{', 'Invalid ArrayObject state'],
        ];
    }

    /** @dataProvider malformedValues */
    #[DataProvider('malformedValues')]
    public function testRejectsMalformedGrammar(string $data, string $message): void
    {
        $this->expectException(UnexpectedValueException::class);
        $this->expectExceptionMessage($message);
        (new DataValidator($data, [ArrayObject::class], 100, 10))->validate();
    }

    /** @return array<string, array{string}> */
    public static function unsafeArrayObjectStates(): array
    {
        return [
            'wrong flag key' => ['i:1;i:0;i:1;a:0:{}i:2;a:0:{}i:3;N;'],
            'non-integer flags' => ['i:0;s:1:"0";i:1;a:0:{}i:2;a:0:{}i:3;N;'],
            'object-backed storage' => ['i:0;i:0;i:1;O:11:"ArrayObject":4:{i:0;i:0;i:1;a:0:{}i:2;a:0:{}i:3;N;}i:2;a:0:{}i:3;N;'],
            'null storage' => ['i:0;i:0;i:1;N;i:2;a:0:{}i:3;N;'],
            'null properties' => ['i:0;i:0;i:1;a:0:{}i:2;N;i:3;N;'],
            'scalar iterator' => ['i:0;i:0;i:1;a:0:{}i:2;a:0:{}i:3;i:0;'],
            'custom iterator' => ['i:0;i:0;i:1;a:0:{}i:2;a:0:{}i:3;s:12:"EvilIterator";'],
            'reordered keys' => ['i:0;i:0;i:2;a:0:{}i:1;a:0:{}i:3;N;'],
        ];
    }

    /** @dataProvider unsafeArrayObjectStates */
    #[DataProvider('unsafeArrayObjectStates')]
    public function testRejectsUnsafeArrayObjectState(string $state): void
    {
        $this->expectException(UnexpectedValueException::class);
        $this->expectExceptionMessage('Unsafe ArrayObject state');
        (new DataValidator('O:11:"ArrayObject":4:{' . $state . '}', [ArrayObject::class], 100, 10))->validate();
    }

    public function testAcceptsExactValueAndDepthBudget(): void
    {
        $this->expectNotToPerformAssertions();
        (new DataValidator('a:1:{i:0;N;}', [], 3, 1))->validate();
    }

    public function testRejectsExhaustedValueBudget(): void
    {
        $this->expectException(UnexpectedValueException::class);
        $this->expectExceptionMessage('Serialized data exceeds resource limits');
        (new DataValidator('N;', [], 0, 1))->validate();
    }

    public function testRejectsDeclaredCollectionOverBudget(): void
    {
        $this->expectException(UnexpectedValueException::class);
        $this->expectExceptionMessage('Serialized collection exceeds resource limits');
        (new DataValidator('a:1:{i:0;N;}', [], 2, 1))->validate();
    }

    public function testRejectsDepthOnePastBudget(): void
    {
        $this->expectException(UnexpectedValueException::class);
        $this->expectExceptionMessage('Serialized data exceeds resource limits');
        (new DataValidator('a:1:{i:0;N;}', [], 3, 0))->validate();
    }
}
