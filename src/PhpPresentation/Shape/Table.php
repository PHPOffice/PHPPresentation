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

namespace PhpOffice\PhpPresentation\Shape;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Shape\Table\Row;

/**
 * Table shape
 */
class Table extends AbstractGraphic implements ComparableInterface
{
    /**
     * Rows
     *
     * @var \PhpOffice\PhpPresentation\Shape\Table\Row[]
     */
    private $rows;

    /**
     * Number of columns
     *
     * @var int
     */
    private $columnCount = 1;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Table instance
     *
     * @param int $columns Number of columns
     */
    public function __construct($columns = 1)
    {
        // Initialise variables
        $this->rows        = array();
        $this->columnCount = $columns;

        // Initialize parent
        parent::__construct();

        // No resize proportional
        $this->resizeProportional = false;
    }

    /**
     * Get row
     *
     * @param  int $row Row number
     * @param  boolean $exceptionAsNull Return a null value instead of an exception?
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\Table\Row
     */
    public function getRow($row = 0, $exceptionAsNull = false)
    {
        if (!isset($this->rows[$row])) {
            if ($exceptionAsNull) {
                return null;
            }
            throw new \Exception('Row number out of bounds.');
        }

        return $this->rows[$row];
    }

    /**
     * Get rows
     *
     * @return \PhpOffice\PhpPresentation\Shape\Table\Row[]
     */
    public function getRows()
    {
        return $this->rows;
    }

    /**
     * Create row
     *
     * @return \PhpOffice\PhpPresentation\Shape\Table\Row
     */
    public function createRow()
    {
        $row           = new Row($this->columnCount);
        $this->rows[] = $row;

        return $row;
    }

    /**
     * @return int
     */
    public function getNumColumns()
    {
        return $this->columnCount;
    }

    /**
     * @param int $numColumn
     * @return Table
     */
    public function setNumColumns($numColumn)
    {
        $this->columnCount = $numColumn;
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        $hashElements = '';
        foreach ($this->rows as $row) {
            $hashElements .= $row->getHashCode();
        }

        return md5($hashElements . __CLASS__);
    }
}
