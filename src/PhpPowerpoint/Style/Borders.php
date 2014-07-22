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

use PhpOffice\PhpPowerpoint\ComparableInterface;

/**
 * \PhpOffice\PhpPowerpoint\Style\Borders
 */
class Borders implements ComparableInterface
{
    /**
     * Left
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Border
     */
    private $left;

    /**
     * Right
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Border
     */
    private $right;

    /**
     * Top
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Border
     */
    private $top;

    /**
     * Bottom
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Border
     */
    private $bottom;

    /**
     * Diagonal up
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Border
     */
    private $diagonalUp;

    /**
     * Diagonal down
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Border
     */
    private $diagonalDown;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Style\Borders
     */
    public function __construct()
    {
        // Initialise values
        $this->left                = new Border();
        $this->right               = new Border();
        $this->top                 = new Border();
        $this->bottom              = new Border();
        $this->diagonalUp          = new Border();
        $this->diagonalUp->setLineStyle(Border::LINE_NONE);
        $this->diagonalDown        = new Border();
        $this->diagonalDown->setLineStyle(Border::LINE_NONE);
    }

    /**
     * Get Left
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Border
     */
    public function getLeft()
    {
        return $this->left;
    }

    /**
     * Get Right
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Border
     */
    public function getRight()
    {
        return $this->right;
    }

    /**
     * Get Top
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Border
     */
    public function getTop()
    {
        return $this->top;
    }

    /**
     * Get Bottom
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Border
     */
    public function getBottom()
    {
        return $this->bottom;
    }

    /**
     * Get Diagonal Up
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Border
     */
    public function getDiagonalUp()
    {
        return $this->diagonalUp;
    }

    /**
     * Get Diagonal Down
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Border
     */
    public function getDiagonalDown()
    {
        return $this->diagonalDown;
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
