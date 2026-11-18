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
 * One stop of a gradient: the colour the gradient is at, and how far along it is.
 *
 * The position is the one `a:gs/@pos` counts, from 0 to 100000, which is a percentage of
 * the gradient with two decimal places' worth of room in it. It is not a percentage, and a
 * position written as one has to be scaled before it is written.
 */
class GradientStop implements ComparableInterface
{
    /** The position of the first stop of a gradient. */
    public const POSITION_START = 0;

    /** The position of the last stop of a gradient. */
    public const POSITION_END = 100000;

    /**
     * Position of the stop along the gradient.
     *
     * @var int
     */
    private $position;

    /**
     * Colour of the stop.
     *
     * @var Color
     */
    private $color;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    public function __construct(int $position = self::POSITION_START, ?Color $color = null)
    {
        $this->position = $position;
        $this->color = $color ?? new Color(Color::COLOR_BLACK);
    }

    /**
     * Get Position.
     */
    public function getPosition(): int
    {
        return $this->position;
    }

    /**
     * Set Position.
     */
    public function setPosition(int $position): self
    {
        $this->position = $position;

        return $this;
    }

    /**
     * Get Color.
     */
    public function getColor(): Color
    {
        return $this->color;
    }

    /**
     * Set Color.
     */
    public function setColor(Color $color): self
    {
        $this->color = $color;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5(
            $this->getPosition()
            . $this->getColor()->getHashCode()
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
