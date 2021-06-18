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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Style;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\Style\Border.
 */
class Border implements ComparableInterface
{
    /* Line style */
    public const LINE_NONE = 'none';
    public const LINE_SINGLE = 'sng';
    public const LINE_DOUBLE = 'dbl';
    public const LINE_THICKTHIN = 'thickThin';
    public const LINE_THINTHICK = 'thinThick';
    public const LINE_TRI = 'tri';

    /* Dash style */
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
     * @var int
     */
    private $lineWidth = 1;

    /**
     * Line style.
     *
     * @var string
     */
    private $lineStyle;

    /**
     * Dash style.
     *
     * @var string
     */
    private $dashStyle;

    /**
     * Border color.
     *
     * @var \PhpOffice\PhpPresentation\Style\Color
     */
    private $color;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Border.
     */
    public function __construct()
    {
        // Initialise values
        $this->lineWidth = 1;
        $this->lineStyle = self::LINE_SINGLE;
        $this->dashStyle = self::DASH_SOLID;
        $this->color = new Color(Color::COLOR_BLACK);
    }

    /**
     * Get line width (in points).
     *
     * @return int
     */
    public function getLineWidth()
    {
        return $this->lineWidth;
    }

    /**
     * Set line width (in points).
     *
     * @param int $pValue
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function setLineWidth($pValue = 1)
    {
        $this->lineWidth = $pValue;

        return $this;
    }

    /**
     * Get line style.
     *
     * @return string
     */
    public function getLineStyle()
    {
        return $this->lineStyle;
    }

    /**
     * Set line style.
     *
     * @param string $pValue
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function setLineStyle($pValue = self::LINE_SINGLE)
    {
        if ('' == $pValue) {
            $pValue = self::LINE_SINGLE;
        }
        $this->lineStyle = $pValue;

        return $this;
    }

    /**
     * Get dash style.
     *
     * @return string
     */
    public function getDashStyle()
    {
        return $this->dashStyle;
    }

    /**
     * Set dash style.
     *
     * @param string $pValue
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function setDashStyle($pValue = self::DASH_SOLID)
    {
        if ('' == $pValue) {
            $pValue = self::DASH_SOLID;
        }
        $this->dashStyle = $pValue;

        return $this;
    }

    /**
     * Get Border Color.
     *
     * @return \PhpOffice\PhpPresentation\Style\Color
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * Set Border Color.
     *
     * @param \PhpOffice\PhpPresentation\Style\Color $color
     *
     * @throws \Exception
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function setColor(Color $color = null)
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
     * @return int|null Hash index
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
