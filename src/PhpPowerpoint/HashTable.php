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

use PhpOffice\PhpPowerpoint\IComparable;

/**
 * PHPPowerPoint_HashTable
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class HashTable
{
    /**
     * HashTable elements
     *
     * @var array
     */
    public $_items = array();

    /**
     * HashTable key map
     *
     * @var array
     */
    public $_keyMap = array();

    /**
     * Create a new PHPPowerPoint_HashTable
     *
     * @param  PhpOffice\PhpPowerpoint\IComparable[] $pSource Optional source array to create HashTable from
     * @throws \Exception
     */
    public function __construct($pSource = null)
    {
        if (!is_null($pSource)) {
            // Create HashTable
            $this->addFromSource($pSource);
        }
    }

    /**
     * Add HashTable items from source
     *
     * @param  PhpOffice\PhpPowerpoint\IComparable[] $pSource Source array to create HashTable from
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
     * @param  PhpOffice\PhpPowerpoint\IComparable $pSource Item to add
     * @throws \Exception
     */
    public function add(IComparable $pSource = null)
    {
        // Determine hashcode
        $hashCode  = null;
        $hashIndex = $pSource->getHashIndex();
        if (is_null($hashIndex)) {
            $hashCode = $pSource->getHashCode();
        } elseif (isset($this->_keyMap[$hashIndex])) {
            $hashCode = $this->_keyMap[$hashIndex];
        } else {
            $hashCode = $pSource->getHashCode();
        }

        // Add value
        if (!isset($this->_items[$hashCode])) {
            $this->_items[$hashCode] = $pSource;
            $index                   = count($this->_items) - 1;
            $this->_keyMap[$index]   = $hashCode;
            $pSource->setHashIndex($index);
        } else {
            $pSource->setHashIndex($this->_items[$hashCode]->getHashIndex());
        }
    }

    /**
     * Remove HashTable item
     *
     * @param  PhpOffice\PhpPowerpoint\IComparable $pSource Item to remove
     * @throws \Exception
     */
    public function remove(IComparable $pSource = null)
    {
        if (isset($this->_items[$pSource->getHashCode()])) {
            unset($this->_items[$pSource->getHashCode()]);

            $deleteKey = -1;
            foreach ($this->_keyMap as $key => $value) {
                if ($deleteKey >= 0) {
                    $this->_keyMap[$key - 1] = $value;
                }

                if ($value == $pSource->getHashCode()) {
                    $deleteKey = $key;
                }
            }
            unset($this->_keyMap[count($this->_keyMap) - 1]);
        }
    }

    /**
     * Clear HashTable
     *
     */
    public function clear()
    {
        $this->_items  = array();
        $this->_keyMap = array();
    }

    /**
     * Count
     *
     * @return int
     */
    public function count()
    {
        return count($this->_items);
    }

    /**
     * Get index for hash code
     *
     * @param  string $pHashCode
     * @return int    Index
     */
    public function getIndexForHashCode($pHashCode = '')
    {
        return array_search($pHashCode, $this->_keyMap);
    }

    /**
     * Get by index
     *
     * @param  int                       $pIndex
     * @return PhpOffice\PhpPowerpoint\IComparable
     *
     */
    public function getByIndex($pIndex = 0)
    {
        if (isset($this->_keyMap[$pIndex])) {
            return $this->getByHashCode($this->_keyMap[$pIndex]);
        }

        return null;
    }

    /**
     * Get by hashcode
     *
     * @param  string                    $pHashCode
     * @return PhpOffice\PhpPowerpoint\IComparable
     *
     */
    public function getByHashCode($pHashCode = '')
    {
        if (isset($this->_items[$pHashCode])) {
            return $this->_items[$pHashCode];
        }

        return null;
    }

    /**
     * HashTable to array
     *
     * @return PhpOffice\PhpPowerpoint\IComparable[]
     */
    public function toArray()
    {
        return $this->_items;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            }
        }
    }
}
