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

namespace PhpOffice\PhpPowerpoint\Shape;

/**
 * PHPPowerPoint_Shape_Hyperlink
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Hyperlink
{
    /**
     * URL to link the shape to
     *
     * @var string
     */
    private $_url;

    /**
     * Tooltip to display on the hyperlink
     *
     * @var string
     */
    private $_tooltip;

    /**
     * Slide number to link to
     *
     * @var int
     */
    private $_slideNumber = null;

    /**
     * Slide relation ID (should not be used by user code!)
     *
     * @var string
     */
    public $__relationId = null;

    /**
     * Create a new PHPPowerPoint_Shape_Hyperlink
     *
     * @param  string    $pUrl     Url to link the shape to
     * @param  string    $pTooltip Tooltip to display on the hyperlink
     * @throws Exception
     */
    public function __construct($pUrl = '', $pTooltip = '')
    {
        // Initialise member variables
        $this->_url     = $pUrl;
        $this->_tooltip = $pTooltip;
    }

    /**
     * Get URL
     *
     * @return string
     */
    public function getUrl()
    {
        return $this->_url;
    }

    /**
     * Set URL
     *
     * @param  string                        $value
     * @return PHPPowerPoint_Shape_Hyperlink
     */
    public function setUrl($value = '')
    {
        $this->_url = $value;

        return $this;
    }

    /**
     * Get slide number
     *
     * @return int
     */
    public function getSlideNumber()
    {
        return $this->_slideNumber;
    }

    /**
     * Set slide number
     *
     * @param  int                           $value
     * @return PHPPowerPoint_Shape_Hyperlink
     */
    public function setSlideNumber($value = 1)
    {
        $this->_url         = 'ppaction://hlinksldjump';
        $this->_slideNumber = $value;

        return $this;
    }

    /**
     * Get tooltip
     *
     * @return string
     */
    public function getTooltip()
    {
        return $this->_tooltip;
    }

    /**
     * Set tooltip
     *
     * @param  string                        $value
     * @return PHPPowerPoint_Shape_Hyperlink
     */
    public function setTooltip($value = '')
    {
        $this->_tooltip = $value;

        return $this;
    }

    /**
     * Is this hyperlink internal? (to another slide)
     *
     * @return boolean
     */
    public function isInternal()
    {
        return strpos($this->_url, 'ppaction://') !== false;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->_url . $this->_tooltip . __CLASS__);
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
}
