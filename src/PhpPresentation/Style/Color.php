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
 * \PhpOffice\PhpPresentation\Style\Color
 */
class Color implements ComparableInterface
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
    private $argb;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Color
     *
     * @param string $pARGB
     */
    public function __construct($pARGB = self::COLOR_BLACK)
    {
        // Initialise values
        $this->argb            = $pARGB;
    }

    /**
     * Get ARGB
     *
     * @return string
     */
    public function getARGB()
    {
        return $this->argb;
    }

    /**
     * Set ARGB
     *
     * @param  string                    $pValue
     * @return \PhpOffice\PhpPresentation\Style\Color
     */
    public function setARGB($pValue = self::COLOR_BLACK)
    {
        if ($pValue == '') {
            $pValue = self::COLOR_BLACK;
        }
        $this->argb = $pValue;

        return $this;
    }

    /**
     * Get the alpha % of the ARGB
     * Will return 100 if no ARGB
     * @return integer
     */
    public function getAlpha()
    {
        $alpha = 100;
        if (strlen($this->argb) >= 6) {
            $dec = hexdec(substr($this->argb, 0, 2));
            $alpha = number_format(($dec/255) * 100, 2);
        }
        return $alpha;
    }

    /**
     * Get RGB
     *
     * @return string
     */
    public function getRGB()
    {
        if (strlen($this->argb) == 6) {
            return $this->argb;
        } else {
            return substr($this->argb, 2);
        }
    }

    /**
     * Set RGB
     *
     * @param  string                    $pValue
     * @return \PhpOffice\PhpPresentation\Style\Color
     */
    public function setRGB($pValue = '000000')
    {
        if ($pValue == '') {
            $pValue = '000000';
        }
        $this->argb = 'FF' . $pValue;

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
            $this->argb
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
