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

namespace PhpOffice\PhpPresentation;

use PhpOffice\Common\Drawing;

/**
 * \PhpOffice\PhpPresentation\DocumentLayout
 */
class DocumentLayout
{
    const LAYOUT_CUSTOM = '';
    const LAYOUT_SCREEN_4X3 = 'screen4x3';
    const LAYOUT_SCREEN_16X10 = 'screen16x10';
    const LAYOUT_SCREEN_16X9 = 'screen16x9';
    const LAYOUT_35MM = '35mm';
    const LAYOUT_A3 = 'A3';
    const LAYOUT_A4 = 'A4';
    const LAYOUT_B4ISO = 'B4ISO';
    const LAYOUT_B5ISO = 'B5ISO';
    const LAYOUT_BANNER = 'banner';
    const LAYOUT_LETTER = 'letter';
    const LAYOUT_OVERHEAD = 'overhead';

    const UNIT_EMU = 'emu';
    const UNIT_CENTIMETER = 'cm';
    const UNIT_INCH = 'in';
    const UNIT_MILLIMETER = 'mm';
    const UNIT_PIXEL = 'px';
    const UNIT_POINT = 'pt';

    /**
     * Dimension types
     *
     * 1 px = 9525 EMU @ 96dpi (which is seems to be the default)
     * Absolute distances are specified in English Metric Units (EMUs),
     * occasionally referred to as A units; there are 360000 EMUs per
     * centimeter, 914400 EMUs per inch, 12700 EMUs per point.
     */
    private $dimension = array(
        self::LAYOUT_SCREEN_4X3 => array('cx' => 9144000, 'cy' => 6858000),
        self::LAYOUT_SCREEN_16X10 => array('cx' => 9144000, 'cy' => 5715000),
        self::LAYOUT_SCREEN_16X9 => array('cx' => 9144000, 'cy' => 5143500),
        self::LAYOUT_35MM => array('cx' => 10287000, 'cy' => 6858000),
        self::LAYOUT_A3 => array('cx' => 15120000, 'cy' => 10692000),
        self::LAYOUT_A4 => array('cx' => 10692000, 'cy' => 7560000),
        self::LAYOUT_B4ISO => array('cx' => 10826750, 'cy' => 8120063),
        self::LAYOUT_B5ISO => array('cx' => 7169150, 'cy' => 5376863),
        self::LAYOUT_BANNER => array('cx' => 7315200, 'cy' => 914400),
        self::LAYOUT_LETTER => array('cx' => 9144000, 'cy' => 6858000),
        self::LAYOUT_OVERHEAD => array('cx' => 9144000, 'cy' => 6858000),
    );

    /**
     * Layout name
     *
     * @var string
     */
    private $layout;

    /**
     * Layout X dimension
     * @var float
     */
    private $dimensionX;

    /**
     * Layout Y dimension
     * @var float
     */
    private $dimensionY;

    /**
     * Create a new \PhpOffice\PhpPresentation\DocumentLayout
     */
    public function __construct()
    {
        $this->setDocumentLayout(self::LAYOUT_SCREEN_4X3);
    }

    /**
     * Get Document Layout
     *
     * @return string
     */
    public function getDocumentLayout()
    {
        return $this->layout;
    }

    /**
     * Set Document Layout
     *
     * @param array|string $pValue
     * @param  boolean $isLandscape
     * @return \PhpOffice\PhpPresentation\DocumentLayout
     */
    public function setDocumentLayout($pValue = self::LAYOUT_SCREEN_4X3, $isLandscape = true)
    {
        switch ($pValue) {
            case self::LAYOUT_SCREEN_4X3:
            case self::LAYOUT_SCREEN_16X10:
            case self::LAYOUT_SCREEN_16X9:
            case self::LAYOUT_35MM:
            case self::LAYOUT_A3:
            case self::LAYOUT_A4:
            case self::LAYOUT_B4ISO:
            case self::LAYOUT_B5ISO:
            case self::LAYOUT_BANNER:
            case self::LAYOUT_LETTER:
            case self::LAYOUT_OVERHEAD:
                $this->layout = $pValue;
                $this->dimensionX = $this->dimension[$this->layout]['cx'];
                $this->dimensionY = $this->dimension[$this->layout]['cy'];
                break;
            case self::LAYOUT_CUSTOM:
            default:
                $this->layout = self::LAYOUT_CUSTOM;
                $this->dimensionX = $pValue['cx'];
                $this->dimensionY = $pValue['cy'];
                break;
        }

        if (!$isLandscape) {
            $tmp = $this->dimensionX;
            $this->dimensionX = $this->dimensionY;
            $this->dimensionY = $tmp;
        }

        return $this;
    }

    /**
     * Get Document Layout cx
     *
     * @param string $unit
     * @return integer
     */
    public function getCX($unit = self::UNIT_EMU)
    {
        return $this->convertUnit($this->dimensionX, self::UNIT_EMU, $unit);
    }

    /**
     * Get Document Layout cy
     *
     * @param string $unit
     * @return integer
     */
    public function getCY($unit = self::UNIT_EMU)
    {
        return $this->convertUnit($this->dimensionY, self::UNIT_EMU, $unit);
    }

    /**
     * Get Document Layout cx
     *
     * @param float $value
     * @param string $unit
     * @return DocumentLayout
     */
    public function setCX($value, $unit = self::UNIT_EMU)
    {
        $this->layout = self::LAYOUT_CUSTOM;
        $this->dimensionX = $this->convertUnit($value, $unit, self::UNIT_EMU);
        return $this;
    }

    /**
     * Get Document Layout cy
     *
     * @param float $value
     * @param string $unit
     * @return DocumentLayout
     */
    public function setCY($value, $unit = self::UNIT_EMU)
    {
        $this->layout = self::LAYOUT_CUSTOM;
        $this->dimensionY = $this->convertUnit($value, $unit, self::UNIT_EMU);
        return $this;
    }

    /**
     * Convert EMUs to differents units
     * @param float $value
     * @param string $fromUnit
     * @param string $toUnit
     * @return float
     */
    protected function convertUnit($value, $fromUnit, $toUnit)
    {
        // Convert from $fromUnit to EMU
        switch ($fromUnit) {
            case self::UNIT_MILLIMETER:
                $value *= 36000;
                break;
            case self::UNIT_CENTIMETER:
                $value *= 360000;
                break;
            case self::UNIT_INCH:
                $value *= 914400;
                break;
            case self::UNIT_PIXEL:
                $value = Drawing::pixelsToEmu($value);
                break;
            case self::UNIT_POINT:
                $value *= 12700;
                break;
            case self::UNIT_EMU:
            default:
                // no changes
        }

        // Convert from EMU to $toUnit
        switch ($toUnit) {
            case self::UNIT_MILLIMETER:
                $value /= 36000;
                break;
            case self::UNIT_CENTIMETER:
                $value /= 360000;
                break;
            case self::UNIT_INCH:
                $value /= 914400;
                break;
            case self::UNIT_PIXEL:
                $value = Drawing::emuToPixels($value);
                break;
            case self::UNIT_POINT:
                $value /= 12700;
                break;
            case self::UNIT_EMU:
            default:
            // no changes
        }
        return $value;
    }
}
