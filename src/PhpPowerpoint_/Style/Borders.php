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

namespace PhpOffice\PhpPowerpoint\Style;

use PhpOffice\PhpPowerpoint\IComparable;
use PhpOffice\PhpPowerpoint\Style\Border;

/**
 * PHPPowerPoint_Style_Borders
 */
class Borders implements IComparable
{
    /**
     * Left
     *
     * @var PHPPowerPoint_Style_Border
     */
    private $_left;

    /**
     * Right
     *
     * @var PHPPowerPoint_Style_Border
     */
    private $_right;

    /**
     * Top
     *
     * @var PHPPowerPoint_Style_Border
     */
    private $_top;

    /**
     * Bottom
     *
     * @var PHPPowerPoint_Style_Border
     */
    private $_bottom;

    /**
     * Diagonal up
     *
     * @var PHPPowerPoint_Style_Border
     */
    private $_diagonalUp;

    /**
     * Diagonal down
     *
     * @var PHPPowerPoint_Style_Border
     */
    private $_diagonalDown;

    /**
     * Create a new PHPPowerPoint_Style_Borders
     */
    public function __construct()
    {
        // Initialise values
        $this->_left                = new Border();
        $this->_right               = new Border();
        $this->_top                 = new Border();
        $this->_bottom              = new Border();
        $this->_diagonalUp          = new Border();
        $this->_diagonalUp->setLineStyle(Border::LINE_NONE);
        $this->_diagonalDown        = new Border();
        $this->_diagonalDown->setLineStyle(Border::LINE_NONE);
    }

    /**
     * Get Left
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getLeft()
    {
        return $this->_left;
    }

    /**
     * Get Right
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getRight()
    {
        return $this->_right;
    }

    /**
     * Get Top
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getTop()
    {
        return $this->_top;
    }

    /**
     * Get Bottom
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getBottom()
    {
        return $this->_bottom;
    }

    /**
     * Get Diagonal Up
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getDiagonalUp()
    {
        return $this->_diagonalUp;
    }

    /**
     * Get Diagonal Down
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getDiagonalDown()
    {
        return $this->_diagonalDown;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5(
            $this->getLeft()->getHashCode()
            . $this->getRight()->getHashCode()
            . $this->getTop()->getHashCode()
            . $this->getBottom()->getHashCode()
            . $this->getDiagonalUp()->getHashCode()
            . $this->getDiagonalDown()->getHashCode()
            . __CLASS__
        );
    }

    /**
     * Hash index
     *
     * @var string
     */
    private $_hashIndex;

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
        return $this->_hashIndex;
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
        $this->_hashIndex = $value;
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
            } else {
                $this->$key = $value;
            }
        }
    }
}
