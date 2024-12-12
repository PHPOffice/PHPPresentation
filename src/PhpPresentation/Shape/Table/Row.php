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

namespace PhpOffice\PhpPresentation\Shape\Table;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\Style\Fill;

/**
 * Table row.
 */
class Row implements ComparableInterface
{
    /**
     * Cells.
     *
     * @var Cell[]
     */
    private $cells = [];

    /**
     * Fill.
     *
     * @var Fill
     */
    private $fill;

    /**
     * Height (in pixels).
     *
     * @var int
     */
    private $height = 38;

    /**
     * Active cell index.
     *
     * @var int
     */
    private $activeCellIndex = -1;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * @param int $columns Number of columns
     */
    public function __construct(int $columns = 1)
    {
        // Fill
        $this->fill = new Fill();
        // Cells
        for ($inc = 0; $inc < $columns; ++$inc) {
            $this->cells[] = new Cell();
        }
    }

    /**
     * Get cell.
     *
     * @param int $cell Cell number
     */
    public function getCell(int $cell = 0): Cell
    {
        if (!isset($this->cells[$cell])) {
            throw new OutOfBoundsException(
                0,
                (count($this->cells) - 1) < 0 ? count($this->cells) - 1 : 0,
                $cell
            );
        }

        return $this->cells[$cell];
    }

    /**
     * Get cell.
     *
     * @param int $cell Cell number
     */
    public function hasCell(int $cell): bool
    {
        return isset($this->cells[$cell]);
    }

    /**
     * Get cells.
     *
     * @return array<Cell>
     */
    public function getCells(): array
    {
        return $this->cells;
    }

    /**
     * Next cell (moves one cell to the right).
     */
    public function nextCell(): Cell
    {
        ++$this->activeCellIndex;
        if (isset($this->cells[$this->activeCellIndex])) {
            $this->cells[$this->activeCellIndex]->setFill(clone $this->getFill());

            return $this->cells[$this->activeCellIndex];
        }

        throw new OutOfBoundsException(
            0,
            (count($this->cells) - 1) < 0 ? count($this->cells) - 1 : 0,
            $this->activeCellIndex
        );
    }

    /**
     * Get fill.
     */
    public function getFill(): Fill
    {
        return $this->fill;
    }

    /**
     * Set fill.
     */
    public function setFill(Fill $fill): self
    {
        $this->fill = $fill;

        return $this;
    }

    /**
     * Get height.
     */
    public function getHeight(): int
    {
        return $this->height;
    }

    /**
     * Set height.
     */
    public function setHeight(int $value = 0): self
    {
        $this->height = $value;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        $hashElements = '';
        foreach ($this->cells as $cell) {
            $hashElements .= $cell->getHashCode();
        }

        return md5($hashElements . $this->fill->getHashCode() . $this->height . __CLASS__);
    }

    /**
     * Get hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return null|int Hash index
     */
    public function getHashIndex(): ?int
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param int $value Hash index
     *
     * @return $this
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;

        return $this;
    }
}
