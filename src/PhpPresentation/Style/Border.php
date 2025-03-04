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

class Border implements ComparableInterface
{
    // Line style
    public const LINE_NONE = 'none';
    public const LINE_SINGLE = 'sng';
    public const LINE_DOUBLE = 'dbl';
    public const LINE_THICKTHIN = 'thickThin';
    public const LINE_THINTHICK = 'thinThick';
    public const LINE_TRI = 'tri';

    // Dash style
    public const DASH_DASH = 'dash';
    public const DASH_DASHDOT = 'dashDot';
    public const DASH_DOT = 'dot';
    public const DASH_LARGEDASH = 'lgDash';
    public const DASH_LARGEDASHDOT = 'lgDashDot';
    public const DASH_LARGEDASHDOTDOT = 'lgDashDotDot';
    public const DASH_SOLID = 'solid';
    public const DASH_SYSDASH = 'sysDash';
    public const DASH_SYSDASHDOT = 'sysDashDot';
    public const DASH_SYSDASHDOTDOT = 'sysDashDotDot';
    public const DASH_SYSDOT = 'sysDot';

    /**
     * Line width.
     *
     * @var float
     */
    private $lineWidth = 1;

    /**
     * Line style.
     *
     * @var string
     */
    private $lineStyle = self::LINE_SINGLE;

    /**
     * Dash style.
     *
     * @var string
     */
    private $dashStyle = self::DASH_SOLID;

    /**
     * Border color.
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

    public function __construct()
    {
        $this->color = new Color(Color::COLOR_BLACK);
    }

    /**
     * Get line width (in pixels).
     */
    public function getLineWidth(): float
    {
        return $this->lineWidth;
    }

    /**
     * Set line width (in pixels).
     */
    public function setLineWidth(float $pValue = 1): self
    {
        $this->lineWidth = $pValue;

        return $this;
    }

    /**
     * Get line style.
     */
    public function getLineStyle(): string
    {
        return $this->lineStyle;
    }

    /**
     * Set line style.
     */
    public function setLineStyle(string $pValue = self::LINE_SINGLE): self
    {
        if ('' == $pValue) {
            $pValue = self::LINE_SINGLE;
        }
        $this->lineStyle = $pValue;

        return $this;
    }

    /**
     * Get dash style.
     */
    public function getDashStyle(): string
    {
        return $this->dashStyle;
    }

    /**
     * Set dash style.
     */
    public function setDashStyle(string $pValue = self::DASH_SOLID): self
    {
        if ('' == $pValue) {
            $pValue = self::DASH_SOLID;
        }
        $this->dashStyle = $pValue;

        return $this;
    }

    /**
     * Get Border Color.
     */
    public function getColor(): ?Color
    {
        return $this->color;
    }

    /**
     * Set Border Color.
     */
    public function setColor(?Color $color = null): self
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
            $this->lineStyle
            . $this->lineWidth
            . $this->dashStyle
            . $this->color->getHashCode()
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
