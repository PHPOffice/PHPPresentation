<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Shape_Hyperlink
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Hyperlink
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
     * @param 	string				$pUrl		Url to link the shape to
     * @param	string				$pTooltip	Tooltip to display on the hyperlink
     * @throws	Exception
     */
    public function __construct($pUrl = '', $pTooltip = '')
    {
    	// Initialise member variables
		$this->_url 		= $pUrl;
		$this->_tooltip 	= $pTooltip;
    }
	
	/**
	 * Get URL
	 *
	 * @return string
	 */
	public function getUrl() {
		return $this->_url;
	}

	/**
	 * Set URL
	 *
	 * @param	string	$value
	 * @return PHPPowerPoint_Shape_Hyperlink
	 */
	public function setUrl($value = '') {
		$this->_url = $value;
		return $this;
	}
	
	/**
	 * Get slide number
	 *
	 * @return int
	 */
	public function getSlideNumber() {
		return $this->_slideNumber;
	}
	
	/**
	 * Set slide number
	 *
	 * @param	int	$value
	 * @return PHPPowerPoint_Shape_Hyperlink
	 */
	public function setSlideNumber($value = 1) {
		$this->_url = 'ppaction://hlinksldjump';
		$this->_slideNumber = $value;
		return $this;
	}
	
	/**
	 * Get tooltip
	 *
	 * @return string
	 */
	public function getTooltip() {
		return $this->_tooltip;
	}

	/**
	 * Set tooltip
	 *
	 * @param	string	$value
	 * @return PHPPowerPoint_Shape_Hyperlink
	 */
	public function setTooltip($value = '') {
		$this->_tooltip = $value;
		return $this;
	}
	
	/**
	 * Is this hyperlink internal? (to another slide)
	 *
	 * @return boolean
	 */
	public function isInternal() {
		return strpos($this->_url, 'ppaction://') !== false;
	}
	
	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */	
	public function getHashCode() {
    	return md5(
    		  $this->_url
    		. $this->_tooltip
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
	 * @return string	Hash index
	 */
	public function getHashIndex() {
		return $this->_hashIndex;
	}

	/**
	 * Set hash index
	 *
	 * Note that this index may vary during script execution! Only reliable moment is
	 * while doing a write of a workbook and when changes are not allowed.
	 *
	 * @param string	$value	Hash index
	 */
	public function setHashIndex($value) {
		$this->_hashIndex = $value;
	}
}