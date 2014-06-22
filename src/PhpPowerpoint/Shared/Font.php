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

namespace PhpOffice\PhpPowerpoint\Shared;

/**
 * Font
 */
class Font
{
    /**
     * Calculate an (approximate) pixel size, based on a font points size
     *
     * @param  int $fontSizeInPoints Font size (in points)
     * @return int Font size (in pixels)
     */
    public static function fontSizeToPixels($fontSizeInPoints = 12)
    {
        return ((16 / 12) * $fontSizeInPoints);
    }

    /**
     * Calculate an (approximate) pixel size, based on inch size
     *
     * @param  int $sizeInInch Font size (in inch)
     * @return int Size (in pixels)
     */
    public static function inchSizeToPixels($sizeInInch = 1)
    {
        return ($sizeInInch * 96);
    }

    /**
     * Calculate an (approximate) pixel size, based on centimeter size
     *
     * @param  int $sizeInCm Font size (in centimeters)
     * @return int Size (in pixels)
     */
    public static function centimeterSizeToPixels($sizeInCm = 1)
    {
        return ($sizeInCm * 37.795275591);
    }
}
