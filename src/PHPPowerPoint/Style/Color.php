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
use PhpOffice\PhpPowerpoint\Style\Color;

/**
 * PHPPowerPoint_Style_Color
 */
class Color implements IComparable
{
    /* Colors */
    const COLOR_BLACK                       = 'FF000000';
    const COLOR_WHITE                       = 'FFFFFFFF';
    const COLOR_RED                         = 'FFFF0000';
    const COLOR_DARKRED                     = 'FF800000';
    const COLOR_BLUE                        = 'FF0000FF';
    const COLOR_DARKBLUE                    = 'FF000080';
    const COLOR_GREEN                       = 'FF00FF00';
    const COLOR_DARKGREEN                   = 'FF008000';
    const COLOR_YELLOW                      = 'FFFFFF00';
    const COLOR_DARKYELLOW                  = 'FF808000';

    /**
     * ARGB - Alpha RGB
     *
     * @var string
     */
    private $_argb;

    /**
     * Create a new PHPPowerPoint_Style_Color
     *
     * @param string $pARGB
     */
    public function __construct($pARGB = self::COLOR_BLACK)
    {
        // Initialise values
        $this->_argb            = $pARGB;
    }

    /**
     * Get ARGB
     *
     * @return string
     */
    public function getARGB()
    {
        return $this->_argb;
    }

    /**
     * Set ARGB
     *
     * @param  string                    $pValue
     * @return PHPPowerPoint_Style_Color
     */
    public function setARGB($pValue = self::COLOR_BLACK)
    {
        if ($pValue == '') {
            $pValue = self::COLOR_BLACK;
        }
        $this->_argb = $pValue;

        return $this;
    }

    /**
     * Get RGB
     *
     * @return string
     */
    public function getRGB()
    {
        if (strlen($this->_argb) == 6) {
            return $this->_argb;
        } else {
            return substr($this->_argb, 2);
        }
    }

    /**
     * Set RGB
     *
     * @param  string                    $pValue
     * @return PHPPowerPoint_Style_Color
     */
    public function setRGB($pValue = '000000')
    {
        if ($pValue == '') {
            $pValue = '000000';
        }
        $this->_argb = 'FF' . $pValue;

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
            $this->_argb
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
