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
 * \PhpOffice\PhpPowerpoint\Shared\Drawing
 */
class Drawing
{
    const DPI_96 = 96;

    /**
     * Convert pixels to EMU
     *
     * @param  int $pValue Value in pixels
     * @return int Value in EMU
     */
    public static function pixelsToEmu($pValue = 0)
    {
        return round($pValue * 9525);
    }

    /**
     * Convert EMU to pixels
     *
     * @param  int $pValue Value in EMU
     * @return int Value in pixels
     */
    public static function emuToPixels($pValue = 0)
    {
        if ($pValue != 0) {
            return round($pValue / 9525);
        } else {
            return 0;
        }
    }

    /**
     * Convert pixels to points
     *
     * @param  int $pValue Value in pixels
     * @return int Value in points
     */
    public static function pixelsToPoints($pValue = 0)
    {
        return $pValue * 0.67777777;
    }

    /**
     * Convert points width to pixels
     *
     * @param  int $pValue Value in points
     * @return int Value in pixels
     */
    public static function pointsToPixels($pValue = 0)
    {
        if ($pValue != 0) {
            return $pValue * 1.333333333;
        } else {
            return 0;
        }
    }

    /**
     * Convert pixels to centimeters
     *
     * @param  int $pValue Value in pixels
     * @return int Value in centimeters
     */
    public static function pixelsToCentimeters($pValue = 0)
    {
        //return $pValue * 0.028;
        return (($pValue / self::DPI_96) * 2.54);
    }

    /**
     * Convert centimeters width to pixels
     *
     * @param  int $pValue Value in centimeters
     * @return int Value in pixels
     */
    public static function centimetersToPixels($pValue = 0)
    {
        if ($pValue != 0) {
            return ($pValue / 2.54) * self::DPI_96;
        } else {
            return 0;
        }
    }

    /**
     * Convert degrees to angle
     *
     * @param  int $pValue Degrees
     * @return int Angle
     */
    public static function degreesToAngle($pValue = 0)
    {
        return (int) round($pValue * 60000);
    }

    /**
     * Convert angle to degrees
     *
     * @param  int $pValue Angle
     * @return int Degrees
     */
    public static function angleToDegrees($pValue = 0)
    {
        if ($pValue != 0) {
            return round($pValue / 60000);
        } else {
            return 0;
        }
    }
}
