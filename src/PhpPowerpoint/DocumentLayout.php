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

namespace PhpOffice\PhpPowerpoint;

/**
 * \PhpOffice\PhpPowerpoint\DocumentLayout
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
     * Layout x dimension
     *
     * @var integer
     */
    private $dimensionX;

    /**
     * Layout y dimension
     *
     * @var integer
     */
    private $dimensionY;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\DocumentLayout
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
     * @return \PhpOffice\PhpPowerpoint\DocumentLayout
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
     * @return integer
     */
    public function getCX()
    {
        return $this->dimensionX;
    }

    /**
     * Get Document Layout cy
     *
     * @return integer
     */
    public function getCY()
    {
        return $this->dimensionY;
    }

    /**
     * Get Document Layout in millimeters
     *
     * @return integer
     */
    public function getLayoutXmilli()
    {
        return $this->dimensionX / 36000;
    }

    /**
     * Get Document Layout in millimeters
     *
     * @return integer
     */
    public function getLayoutYmilli()
    {
        return $this->dimensionY / 36000;
    }

    /**
     * Set Document Layout in millimeters
     *
     * @param  integer                      $pValue Layout width
     * @return \PhpOffice\PhpPowerpoint\DocumentLayout
     */
    public function setLayoutXmilli($pValue)
    {
        $this->dimensionX     = $pValue * 36000;
        $this->layout = self::LAYOUT_CUSTOM;

        return $this;
    }
    /**
     * Set Document Layout in millimeters
     *
     * @param  integer                      $pValue Layout height
     * @return \PhpOffice\PhpPowerpoint\DocumentLayout
     */
    public function setLayoutYmilli($pValue)
    {
        $this->dimensionY     = $pValue * 36000;
        $this->layout = self::LAYOUT_CUSTOM;

        return $this;
    }
}
