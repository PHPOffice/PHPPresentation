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

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AbstractGraphic;

abstract class AbstractDecoratorWriter
{
    abstract public function render(): ZipInterface;

    /**
     * @var HashTable
     */
    protected $oHashTable;

    /**
     * @var PhpPresentation
     */
    protected $oPresentation;

    /**
     * @var ZipInterface
     */
    protected $oZip;

    /**
     * @return $this
     */
    public function setDrawingHashTable(HashTable $hashTable)
    {
        $this->oHashTable = $hashTable;

        return $this;
    }

    /**
     * @return HashTable
     */
    public function getDrawingHashTable()
    {
        return $this->oHashTable;
    }

    /**
     * The graphic whose bytes were actually written for this one.
     *
     * `HashTable` keeps one entry per hash, so two shapes that compare alike leave a single part
     * in the archive -- the first of them. The others are handed that one's hash index when they
     * are added, and this reads it back. A reference has to name the part that exists, not the
     * name this shape would have had; `getIndexedFilename()` counts instances and knows nothing
     * about the collapse.
     *
     * Answers with the shape itself when the table has not been filled yet, so a writer that runs
     * without one behaves as it did before.
     */
    protected function writtenPart(AbstractGraphic $shape): AbstractGraphic
    {
        $hashIndex = $shape->getHashIndex();
        $written = null === $hashIndex || null === $this->oHashTable
            ? null
            : $this->oHashTable->getByIndex($hashIndex);

        return $written instanceof AbstractGraphic ? $written : $shape;
    }

    /**
     * @return $this
     */
    public function setPresentation(PhpPresentation $oPresentation)
    {
        $this->oPresentation = $oPresentation;

        return $this;
    }

    /**
     * @return PhpPresentation
     */
    public function getPresentation()
    {
        return $this->oPresentation;
    }

    /**
     * @return $this
     */
    public function setZip(ZipInterface $oZip)
    {
        $this->oZip = $oZip;

        return $this;
    }

    /**
     * @return ZipInterface
     */
    public function getZip()
    {
        return $this->oZip;
    }
}
