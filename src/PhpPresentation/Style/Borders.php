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

namespace PhpOffice\PhpPresentation\Style;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\Style\Borders.
 */
class Borders implements ComparableInterface
{
    /**
     * Left.
     *
     * @var Border
     */
    private $left;

    /**
     * Right.
     *
     * @var Border
     */
    private $right;

    /**
     * Top.
     *
     * @var Border
     */
    private $top;

    /**
     * Bottom.
     *
     * @var Border
     */
    private $bottom;

    /**
     * Diagonal up.
     *
     * @var Border
     */
    private $diagonalUp;

    /**
     * Diagonal down.
     *
     * @var Border
     */
    private $diagonalDown;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Borders.
     */
    public function __construct()
    {
        // Initialise values
        $this->left = new Border();
        $this->right = new Border();
        $this->top = new Border();
        $this->bottom = new Border();
        $this->diagonalUp = new Border();
        $this->diagonalUp->setLineStyle(Border::LINE_NONE);
        $this->diagonalDown = new Border();
        $this->diagonalDown->setLineStyle(Border::LINE_NONE);
    }

    /**
     * Get Left.
     *
     * @return Border
     */
    public function getLeft()
    {
        return $this->left;
    }

    /**
     * Get Right.
     *
     * @return Border
     */
    public function getRight()
    {
        return $this->right;
    }

    /**
     * Get Top.
     *
     * @return Border
     */
    public function getTop()
    {
        return $this->top;
    }

    /**
     * Get Bottom.
     *
     * @return Border
     */
    public function getBottom()
    {
        return $this->bottom;
    }

    /**
     * Get Diagonal Up.
     *
     * @return Border
     */
    public function getDiagonalUp()
    {
        return $this->diagonalUp;
    }

    /**
     * Get Diagonal Down.
     *
     * @return Border
     */
    public function getDiagonalDown()
    {
        return $this->diagonalDown;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
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
