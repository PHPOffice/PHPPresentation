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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\HashTable
 */
class HashTable
{
    /**
     * HashTable elements
     *
     * @var array<string, ComparableInterface>
     */
    public $items = array();

    /**
     * HashTable key map
     *
     * @var array<int, string>
     */
    public $keyMap = array();

    /**
     * Create a new \PhpOffice\PhpPresentation\HashTable
     *
     * @param array<int, ComparableInterface> $pSource Optional source array to create HashTable from
     * @throws \Exception
     */
    public function __construct(array $pSource = [])
    {
        $this->addFromSource($pSource);
    }

    /**
     * Add HashTable items from source
     *
     * @param array<int, ComparableInterface> $pSource Source array to create HashTable from
     */
    public function addFromSource(array $pSource = []): void
    {
        foreach ($pSource as $item) {
            $this->add($item);
        }
    }

    /**
     * Add HashTable item
     *
     * @param ComparableInterface $pSource Item to add
     */
    public function add(ComparableInterface $pSource): void
    {
        // Determine hashcode
        $hashIndex = $pSource->getHashIndex();
        $hashCode = $pSource->getHashCode();
        if (isset($this->keyMap[$hashIndex])) {
            $hashCode = $this->keyMap[$hashIndex];
        }

        // Add value
        if (!isset($this->items[$hashCode])) {
            $this->items[$hashCode] = $pSource;
            $index = count($this->items) - 1;
            $this->keyMap[$index] = $hashCode;
            $pSource->setHashIndex($index);
        } else {
            $pSource->setHashIndex($this->items[$hashCode]->getHashIndex());
        }
    }

    /**
     * Remove HashTable item
     *
     * @param ComparableInterface $pSource Item to remove
     * @throws \Exception
     */
    public function remove(ComparableInterface $pSource): void
    {
        if (isset($this->items[$pSource->getHashCode()])) {
            unset($this->items[$pSource->getHashCode()]);

            $deleteKey = -1;
            foreach ($this->keyMap as $key => $value) {
                if ($deleteKey >= 0) {
                    $this->keyMap[$key - 1] = $value;
                }

                if ($value == $pSource->getHashCode()) {
                    $deleteKey = $key;
                }
            }
            unset($this->keyMap[count($this->keyMap) - 1]);
        }
    }

    /**
     * Clear HashTable
     */
    public function clear(): void
    {
        $this->items  = array();
        $this->keyMap = array();
    }

    /**
     * Count
     *
     * @return int
     */
    public function count(): int
    {
        return count($this->items);
    }

    /**
     * Get index for hash code
     *
     * @param string $pHashCode
     * @return int Index (-1 if not found)
     */
    public function getIndexForHashCode(string $pHashCode = ''): int
    {
        $index = array_search($pHashCode, $this->keyMap);
        return $index === false ? -1 : $index;
    }

    /**
     * Get by index
     *
     * @param int $pIndex
     * @return ComparableInterface|null
     *
     */
    public function getByIndex(int $pIndex = 0): ?ComparableInterface
    {
        if (isset($this->keyMap[$pIndex])) {
            return $this->getByHashCode($this->keyMap[$pIndex]);
        }

        return null;
    }

    /**
     * Get by hashcode
     *
     * @param string $pHashCode
     * @return ComparableInterface|null
     *
     */
    public function getByHashCode(string $pHashCode = ''): ?ComparableInterface
    {
        if (isset($this->items[$pHashCode])) {
            return $this->items[$pHashCode];
        }

        return null;
    }

    /**
     * HashTable to array
     *
     * @return array<ComparableInterface>
     */
    public function toArray(): array
    {
        return $this->items;
    }
}
