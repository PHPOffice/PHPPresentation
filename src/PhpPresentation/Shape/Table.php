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

namespace PhpOffice\PhpPresentation\Shape;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\Shape\Table\Row;

/**
 * Table shape.
 */
class Table extends AbstractGraphic implements ComparableInterface
{
    /**
     * Rows.
     *
     * @var array<int, Row>
     */
    private $rows = [];

    /**
     * Number of columns.
     *
     * @var int
     */
    private $columnCount = 1;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Table instance.
     *
     * @param int $columns Number of columns
     */
    public function __construct($columns = 1)
    {
        $this->columnCount = $columns;

        // Initialize parent
        parent::__construct();

        // No resize proportional
        $this->resizeProportional = false;
    }

    /**
     * Get row.
     *
     * @param int $row Row number
     */
    public function getRow(int $row = 0): Row
    {
        if (!isset($this->rows[$row])) {
            throw new OutOfBoundsException(
                0,
                (count($this->rows) - 1) < 0 ? 0 : count($this->rows) - 1,
                $row
            );
        }

        return $this->rows[$row];
    }

    public function hasRow(int $row): bool
    {
        return isset($this->rows[$row]);
    }

    /**
     * Get rows.
     *
     * @return Row[]
     */
    public function getRows(): array
    {
        return $this->rows;
    }

    /**
     * Create row.
     */
    public function createRow(): Row
    {
        $row = new Row($this->columnCount);
        $this->rows[] = $row;

        return $row;
    }

    public function getNumColumns(): int
    {
        return $this->columnCount;
    }

    public function setNumColumns(int $numColumn): self
    {
        $this->columnCount = $numColumn;

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
        foreach ($this->rows as $row) {
            $hashElements .= $row->getHashCode();
        }

        return md5($hashElements . __CLASS__);
    }
}
