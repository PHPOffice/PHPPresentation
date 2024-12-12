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
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;

class Alignment implements ComparableInterface
{
    // Horizontal alignment
    public const HORIZONTAL_GENERAL = 'l';
    public const HORIZONTAL_LEFT = 'l';
    public const HORIZONTAL_RIGHT = 'r';
    public const HORIZONTAL_CENTER = 'ctr';
    public const HORIZONTAL_JUSTIFY = 'just';
    public const HORIZONTAL_DISTRIBUTED = 'dist';

    // Vertical alignment
    public const VERTICAL_BASE = 'base';
    public const VERTICAL_AUTO = 'auto';
    public const VERTICAL_BOTTOM = 'b';
    public const VERTICAL_TOP = 't';
    public const VERTICAL_CENTER = 'ctr';

    // Text direction
    public const TEXT_DIRECTION_HORIZONTAL = 'horz';
    public const TEXT_DIRECTION_VERTICAL_90 = 'vert';
    public const TEXT_DIRECTION_VERTICAL_270 = 'vert270';
    public const TEXT_DIRECTION_STACKED = 'wordArtVert';

    /**
     * @var array<int, string>
     */
    private $supportedStyles = [
        self::HORIZONTAL_GENERAL,
        self::HORIZONTAL_LEFT,
        self::HORIZONTAL_RIGHT,
    ];

    /**
     * Horizontal.
     *
     * @var string
     */
    private $horizontal = self::HORIZONTAL_LEFT;

    /**
     * Vertical.
     *
     * @var string
     */
    private $vertical = self::VERTICAL_BASE;

    /**
     * Text Direction.
     *
     * @var string
     */
    private $textDirection = self::TEXT_DIRECTION_HORIZONTAL;

    /**
     * Level.
     *
     * @var int
     */
    private $level = 0;

    /**
     * Indent - only possible with horizontal alignment left and right.
     *
     * @var float
     */
    private $indent = 0;

    /**
     * Margin left - only possible with horizontal alignment left and right.
     *
     * @var float
     */
    private $marginLeft = 0;

    /**
     * Margin right - only possible with horizontal alignment left and right.
     *
     * @var float
     */
    private $marginRight = 0;

    /**
     * Margin top.
     *
     * @var float
     */
    private $marginTop = 0;

    /**
     * Margin bottom.
     *
     * @var float
     */
    private $marginBottom = 0;

    /**
     * RTL Direction Support.
     *
     * @var bool
     */
    private $isRTL = false;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Get Horizontal.
     */
    public function getHorizontal(): string
    {
        return $this->horizontal;
    }

    /**
     * Set Horizontal.
     */
    public function setHorizontal(string $pValue = self::HORIZONTAL_LEFT): self
    {
        if ('' == $pValue) {
            $pValue = self::HORIZONTAL_LEFT;
        }
        $this->horizontal = $pValue;

        return $this;
    }

    /**
     * Get Vertical.
     */
    public function getVertical(): string
    {
        return $this->vertical;
    }

    /**
     * Set Vertical.
     */
    public function setVertical(string $pValue = self::VERTICAL_BASE): self
    {
        if ('' == $pValue) {
            $pValue = self::VERTICAL_BASE;
        }
        $this->vertical = $pValue;

        return $this;
    }

    /**
     * Get Level.
     */
    public function getLevel(): int
    {
        return $this->level;
    }

    /**
     * Set Level.
     *
     * @param int $pValue Ranging 0 - 8
     */
    public function setLevel(int $pValue = 0): self
    {
        if ($pValue < 0) {
            throw new OutOfBoundsException(0, null, $pValue);
        }
        $this->level = $pValue;

        return $this;
    }

    /**
     * Get indent.
     */
    public function getIndent(): float
    {
        return $this->indent;
    }

    /**
     * Set indent.
     */
    public function setIndent(float $pValue = 0): self
    {
        if ($pValue > 0 && !in_array($this->getHorizontal(), $this->supportedStyles)) {
            $pValue = 0; // indent not supported
        }

        $this->indent = $pValue;

        return $this;
    }

    /**
     * Get margin left.
     */
    public function getMarginLeft(): float
    {
        return $this->marginLeft;
    }

    /**
     * Set margin left.
     */
    public function setMarginLeft(float $pValue = 0): self
    {
        if ($pValue > 0 && !in_array($this->getHorizontal(), $this->supportedStyles)) {
            $pValue = 0; // margin left not supported
        }

        $this->marginLeft = $pValue;

        return $this;
    }

    /**
     * Get margin right.
     */
    public function getMarginRight(): float
    {
        return $this->marginRight;
    }

    /**
     * Set margin ight.
     */
    public function setMarginRight(float $pValue = 0): self
    {
        if ($pValue > 0 && !in_array($this->getHorizontal(), $this->supportedStyles)) {
            $pValue = 0; // margin right not supported
        }

        $this->marginRight = $pValue;

        return $this;
    }

    /**
     * Get margin top.
     */
    public function getMarginTop(): float
    {
        return $this->marginTop;
    }

    /**
     * Set margin top.
     */
    public function setMarginTop(float $pValue = 0): self
    {
        $this->marginTop = $pValue;

        return $this;
    }

    /**
     * Get margin bottom.
     */
    public function getMarginBottom(): float
    {
        return $this->marginBottom;
    }

    /**
     * Set margin bottom.
     */
    public function setMarginBottom(float $pValue = 0): self
    {
        $this->marginBottom = $pValue;

        return $this;
    }

    public function getTextDirection(): string
    {
        return $this->textDirection;
    }

    public function setTextDirection(string $pValue = self::TEXT_DIRECTION_HORIZONTAL): self
    {
        if (empty($pValue)) {
            $pValue = self::TEXT_DIRECTION_HORIZONTAL;
        }
        $this->textDirection = $pValue;

        return $this;
    }

    public function isRTL(): bool
    {
        return $this->isRTL;
    }

    public function setIsRTL(bool $value = false): self
    {
        $this->isRTL = $value;

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
            $this->horizontal
            . $this->vertical
            . $this->level
            . $this->indent
            . $this->marginLeft
            . $this->marginRight
            . ($this->isRTL ? '1' : '0')
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
