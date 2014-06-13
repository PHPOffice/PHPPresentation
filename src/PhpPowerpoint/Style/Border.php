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

use \PhpOffice\PhpPowerpoint\IComparable;
use \PhpOffice\PhpPowerpoint\Style\Color;

/**
 * PHPPowerPoint_Style_Border
 */
class Border implements IComparable
{
    /* Line style */
    const LINE_NONE             = 'none';
    const LINE_SINGLE           = 'sng';
    const LINE_DOUBLE           = 'dbl';
    const LINE_THICKTHIN        = 'thickThin';
    const LINE_THINTHICK        = 'thinThick';
    const LINE_TRI              = 'tri';

    /* Dash style */
    const DASH_DASH             = 'dash';
    const DASH_DASHDOT          = 'dashDot';
    const DASH_DOT              = 'dot';
    const DASH_LARGEDASH        = 'lgDash';
    const DASH_LARGEDASHDOT     = 'lgDashDot';
    const DASH_LARGEDASHDOTDOT  = 'lgDashDotDot';
    const DASH_SOLID            = 'solid';
    const DASH_SYSDASH          = 'sysDash';
    const DASH_SYSDASHDOT       = 'sysDashDot';
    const DASH_SYSDASHDOTDOT    = 'sysDashDotDot';
    const DASH_SYSDOT           = 'sysDot';

    /**
     * Line width
     *
     * @var int
     */
    private $lineWidth = 1;

    /**
     * Line style
     *
     * @var string
     */
    private $lineStyle;

    /**
     * Dash style
     *
     * @var string
     */
    private $dashStyle;

    /**
     * Border color
     *
     * @var PHPPowerPoint_Style_Color
     */
    private $color;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new PHPPowerPoint_Style_Border
     */
    public function __construct()
    {
        // Initialise values
        $this->lineWidth = 1;
        $this->lineStyle = self::LINE_SINGLE;
        $this->dashStyle = self::DASH_SOLID;
        $this->color     = new Color(Color::COLOR_BLACK);
    }

    /**
     * Get line width
     *
     * @return int
     */
    public function getLineWidth()
    {
        return $this->lineWidth;
    }

    /**
     * Set line width
     *
     * @param  int                        $pValue
     * @return PHPPowerPoint_Style_Border
     */
    public function setLineWidth($pValue = 1)
    {
        $this->lineWidth = $pValue;

        return $this;
    }

    /**
     * Get line style
     *
     * @return string
     */
    public function getLineStyle()
    {
        return $this->lineStyle;
    }

    /**
     * Set line style
     *
     * @param  string                     $pValue
     * @return PHPPowerPoint_Style_Border
     */
    public function setLineStyle($pValue = self::LINE_SINGLE)
    {
        if ($pValue == '') {
            $pValue = self::LINE_SINGLE;
        }
        $this->lineStyle = $pValue;

        return $this;
    }

    /**
     * Get dash style
     *
     * @return string
     */
    public function getDashStyle()
    {
        return $this->dashStyle;
    }

    /**
     * Set dash style
     *
     * @param  string                     $pValue
     * @return PHPPowerPoint_Style_Border
     */
    public function setDashStyle($pValue = self::DASH_SOLID)
    {
        if ($pValue == '') {
            $pValue = self::DASH_SOLID;
        }
        $this->dashStyle = $pValue;

        return $this;
    }

    /**
     * Get Border Color
     *
     * @return PHPPowerPoint_Style_Color
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * Set Border Color
     *
     * @param  PHPPowerPoint_Style_Color  $color
     * @throws \Exception
     * @return PHPPowerPoint_Style_Border
     */
    public function setColor(Color $color = null)
    {
        $this->color = $color;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
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
