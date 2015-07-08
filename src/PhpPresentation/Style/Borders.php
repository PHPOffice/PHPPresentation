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

namespace PhpOffice\PhpPresentation\Style;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\Style\Borders
 */
class Borders implements ComparableInterface
{
    /**
     * Left
     *
     * @var \PhpOffice\PhpPresentation\Style\Border
     */
    private $left;

    /**
     * Right
     *
     * @var \PhpOffice\PhpPresentation\Style\Border
     */
    private $right;

    /**
     * Top
     *
     * @var \PhpOffice\PhpPresentation\Style\Border
     */
    private $top;

    /**
     * Bottom
     *
     * @var \PhpOffice\PhpPresentation\Style\Border
     */
    private $bottom;

    /**
     * Diagonal up
     *
     * @var \PhpOffice\PhpPresentation\Style\Border
     */
    private $diagonalUp;

    /**
     * Diagonal down
     *
     * @var \PhpOffice\PhpPresentation\Style\Border
     */
    private $diagonalDown;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Borders
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
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function getLeft()
    {
        return $this->left;
    }

    /**
     * Get Right
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function getRight()
    {
        return $this->right;
    }

    /**
     * Get Top
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function getTop()
    {
        return $this->top;
    }

    /**
     * Get Bottom
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function getBottom()
    {
        return $this->bottom;
    }

    /**
     * Get Diagonal Up
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function getDiagonalUp()
    {
        return $this->diagonalUp;
    }

    /**
     * Get Diagonal Down
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
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
