<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPowerpoint;

use PhpOffice\PhpPowerpoint\ComparableInterface;

/**
 * \PhpOffice\PhpPowerpoint\HashTable
 */
class HashTable
{
    /**
     * HashTable elements
     *
     * @var array
     */
    public $items = array();

    /**
     * HashTable key map
     *
     * @var array
     */
    public $keyMap = array();

    /**
     * Create a new \PhpOffice\PhpPowerpoint\HashTable
     *
     * @param  \PhpOffice\PhpPowerpoint\ComparableInterface[] $pSource Optional source array to create HashTable from
     * @throws \Exception
     */
    public function __construct(array $pSource = null)
    {
        if (!is_null($pSource)) {
            // Create HashTable
            $this->addFromSource($pSource);
        }
    }

    /**
     * Add HashTable items from source
     *
     * @param  \PhpOffice\PhpPowerpoint\ComparableInterface[] $pSource Source array to create HashTable from
     * @throws \Exception
     */
    public function addFromSource($pSource = null)
    {
        // Check if an array was passed
        if ($pSource == null) {
            return;
        } elseif (!is_array($pSource)) {
            throw new \Exception('Invalid array parameter passed.');
        }

        foreach ($pSource as $item) {
            $this->add($item);
        }
    }

    /**
     * Add HashTable item
     *
     * @param \PhpOffice\PhpPowerpoint\ComparableInterface $pSource Item to add
     */
    public function add(ComparableInterface $pSource)
    {
        // Determine hashcode
        $hashIndex = $pSource->getHashIndex();
        if (is_null($hashIndex)) {
            $hashCode = $pSource->getHashCode();
        } elseif (isset($this->keyMap[$hashIndex])) {
            $hashCode = $this->keyMap[$hashIndex];
        } else {
            $hashCode = $pSource->getHashCode();
        }

        // Add value
        if (!isset($this->items[$hashCode])) {
            $this->items[$hashCode] = $pSource;
            $index                   = count($this->items) - 1;
            $this->keyMap[$index]   = $hashCode;
            $pSource->setHashIndex($index);
        } else {
            $pSource->setHashIndex($this->items[$hashCode]->getHashIndex());
        }
    }

    /**
     * Remove HashTable item
     *
     * @param  \PhpOffice\PhpPowerpoint\ComparableInterface $pSource Item to remove
     * @throws \Exception
     */
    public function remove(ComparableInterface $pSource)
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
     *
     */
    public function clear()
    {
        $this->items  = array();
        $this->keyMap = array();
    }

    /**
     * Count
     *
     * @return int
     */
    public function count()
    {
        return count($this->items);
    }

    /**
     * Get index for hash code
     *
     * @param  string $pHashCode
     * @return int    Index
     */
    public function getIndexForHashCode($pHashCode = '')
    {
        return array_search($pHashCode, $this->keyMap);
    }

    /**
     * Get by index
     *
     * @param  int                       $pIndex
     * @return \PhpOffice\PhpPowerpoint\ComparableInterface
     *
     */
    public function getByIndex($pIndex = 0)
    {
        if (isset($this->keyMap[$pIndex])) {
            return $this->getByHashCode($this->keyMap[$pIndex]);
        }

        return null;
    }

    /**
     * Get by hashcode
     *
     * @param  string                    $pHashCode
     * @return \PhpOffice\PhpPowerpoint\ComparableInterface
     *
     */
    public function getByHashCode($pHashCode = '')
    {
        if (isset($this->items[$pHashCode])) {
            return $this->items[$pHashCode];
        }

        return null;
    }

    /**
     * HashTable to array
     *
     * @return \PhpOffice\PhpPowerpoint\ComparableInterface[]
     */
    public function toArray()
    {
        return $this->items;
    }
}
