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

namespace PhpOffice\PhpPresentation\Shape\Table;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Style\Fill;

/**
 * Table row
 */
class Row implements ComparableInterface
{
    /**
     * Cells
     *
     * @var \PhpOffice\PhpPresentation\Shape\Table\Cell[]
     */
    private $cells;

    /**
     * Fill
     *
     * @var \PhpOffice\PhpPresentation\Style\Fill
     */
    private $fill;

    /**
     * Height (in pixels)
     *
     * @var int
     */
    private $height = 38;

    /**
     * Active cell index
     *
     * @var int
     */
    private $activeCellIndex = -1;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Table\Row instance
     *
     * @param int $columns Number of columns
     */
    public function __construct($columns = 1)
    {
        // Initialise variables
        $this->cells = array();
        for ($i = 0; $i < $columns; $i++) {
            $this->cells[] = new Cell();
        }

        // Set fill
        $this->fill = new Fill();
    }

    /**
     * Get cell
     *
     * @param  int $cell Cell number
     * @param  boolean $exceptionAsNull Return a null value instead of an exception?
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\Table\Cell
     */
    public function getCell($cell = 0, $exceptionAsNull = false)
    {
        if (!isset($this->cells[$cell])) {
            if ($exceptionAsNull) {
                return null;
            }
            throw new \Exception('Cell number out of bounds.');
        }

        return $this->cells[$cell];
    }

    /**
     * Get cells
     *
     * @return \PhpOffice\PhpPresentation\Shape\Table\Cell[]
     */
    public function getCells()
    {
        return $this->cells;
    }

    /**
     * Next cell (moves one cell to the right)
     *
     * @return \PhpOffice\PhpPresentation\Shape\Table\Cell
     * @throws \Exception
     */
    public function nextCell()
    {
        $this->activeCellIndex++;
        if (isset($this->cells[$this->activeCellIndex])) {
            $this->cells[$this->activeCellIndex]->setFill(clone $this->getFill());

            return $this->cells[$this->activeCellIndex];
        } else {
            throw new \Exception("Cell count out of bounds.");
        }
    }

    /**
     * Get fill
     *
     * @return \PhpOffice\PhpPresentation\Style\Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Set fill
     *
     * @param  \PhpOffice\PhpPresentation\Style\Fill      $fill
     * @return \PhpOffice\PhpPresentation\Shape\Table\Row
     */
    public function setFill(Fill $fill)
    {
        $this->fill = $fill;

        return $this;
    }

    /**
     * Get height
     *
     * @return int
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Set height
     *
     * @param  int                          $value
     * @return \PhpOffice\PhpPresentation\Shape\RichText
     */
    public function setHeight($value = 0)
    {
        $this->height = $value;

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
        foreach ($this->cells as $cell) {
            $hashElements .= $cell->getHashCode();
        }

        return md5($hashElements . $this->fill->getHashCode() . $this->height . __CLASS__);
    }

    /**
     * Get hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return string Hash index
     */
    public function getHashIndex()
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param string $value Hash index
     */
    public function setHashIndex($value)
    {
        $this->hashIndex = $value;
    }
}
