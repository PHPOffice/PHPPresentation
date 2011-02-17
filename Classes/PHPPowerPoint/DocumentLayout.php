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
 * @package    PHPPowerPoint_Style
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_DocumentLayout
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_DocumentLayout
{
	const LAYOUT_CUSTOM        = '';
	const LAYOUT_SCREEN_4x3    = 'screen4x3';
	const LAYOUT_SCREEN_16x10  = 'screen16x10';
	const LAYOUT_SCREEN_16x9   = 'screen16x9';
	const LAYOUT_35mm          = '35mm';
	const LAYOUT_A3            = 'A3';
	const LAYOUT_A4            = 'A4';
	const LAYOUT_B4ISO         = 'B4ISO';
	const LAYOUT_B5ISO         = 'B5ISO';
	const LAYOUT_BANNER        = 'banner';
	const LAYOUT_LETTER        = 'letter';
	const LAYOUT_OVERHEAD      = 'overhead';

	/*
	 * 1 px = 9525 EMU @ 96dpi (which is seems to be the default)
	 * Absolute distances are specified in English Metric Units (EMUs), 
	 * occasionally referred to as A units; there are 360000 EMUs per 
	 * centimeter, 914400 EMUs per inch, 12700 EMUs per point.
	 */
	private $_dimension = array(
		PHPPowerPoint_DocumentLayout::LAYOUT_SCREEN_4x3   => array('cx' =>  9144000, 'cy' => 6858000),
		PHPPowerPoint_DocumentLayout::LAYOUT_SCREEN_16x10 => array('cx' =>  9144000, 'cy' => 5715000),
		PHPPowerPoint_DocumentLayout::LAYOUT_SCREEN_16x9  => array('cx' =>  9144000, 'cy' => 5143500),
		PHPPowerPoint_DocumentLayout::LAYOUT_35mm         => array('cx' => 10287000, 'cy' => 6858000),
		PHPPowerPoint_DocumentLayout::LAYOUT_A3           => array('cx' => 12801600, 'cy' => 9601200),
		PHPPowerPoint_DocumentLayout::LAYOUT_A4           => array('cx' =>  9906000, 'cy' => 6858000),
		PHPPowerPoint_DocumentLayout::LAYOUT_B4ISO        => array('cx' => 10826750, 'cy' => 8120063),
		PHPPowerPoint_DocumentLayout::LAYOUT_B5ISO        => array('cx' =>  7169150, 'cy' => 5376863),
		PHPPowerPoint_DocumentLayout::LAYOUT_BANNER       => array('cx' =>  7315200, 'cy' =>  914400),
		PHPPowerPoint_DocumentLayout::LAYOUT_LETTER       => array('cx' =>  9144000, 'cy' => 6858000),
		PHPPowerPoint_DocumentLayout::LAYOUT_OVERHEAD     => array('cx' =>  9144000, 'cy' => 6858000)
	);


	/**
	 * Fill type
	 *
	 * @var string
	 */
	private $_layout;
	private $_cx;
	private $_cy;

    /**
     * Create a new PHPPowerPoint_DocumentLayout
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_layout = PHPPowerPoint_DocumentLayout::LAYOUT_SCREEN_4x3;
	$this->_cx = $this->_dimension[$this->_layout]['cx'];
	$this->_cy = $this->_dimension[$this->_layout]['cy'];
    }

    /**
     * Get Document Layout
     *
     * @return string
     */
    public function getDocumentLayout() {
    	return $this->_layout;
    }

    /**
     * Set Document Layout
     *
     * @param string $pValue	PHPPowerPoint_DocumentLayout document layout
     * @return PHPPowerPoint_DocumentLayout
     */
    public function setDocumentLayout($pValue = PHPPowerPoint_DocumentLayout::LAYOUT_SCREEN_4x3, $isLandscape = true) {
	switch($pValue) {
		case PHPPowerPoint_DocumentLayout::LAYOUT_SCREEN_4x3:
		case PHPPowerPoint_DocumentLayout::LAYOUT_SCREEN_16x10:
		case PHPPowerPoint_DocumentLayout::LAYOUT_SCREEN_16x9:
		case PHPPowerPoint_DocumentLayout::LAYOUT_35mm:
		case PHPPowerPoint_DocumentLayout::LAYOUT_A3:
		case PHPPowerPoint_DocumentLayout::LAYOUT_A4:
		case PHPPowerPoint_DocumentLayout::LAYOUT_B4ISO:
		case PHPPowerPoint_DocumentLayout::LAYOUT_B5ISO:
		case PHPPowerPoint_DocumentLayout::LAYOUT_BANNER:
		case PHPPowerPoint_DocumentLayout::LAYOUT_LETTER:
		case PHPPowerPoint_DocumentLayout::LAYOUT_OVERHEAD:
			$this->_layout = $pValue;
			if( $isLandscape ) {
				$this->_cx = $this->_dimension[$this->_layout]['cx'];
				$this->_cy = $this->_dimension[$this->_layout]['cy'];
			} else {
				$this->_cx = $this->_dimension[$this->_layout]['cy'];
				$this->_cy = $this->_dimension[$this->_layout]['cx'];
			}
			break;
		case PHPPowerPoint_DocumentLayout::LAYOUT_CUSTOM:
		default:
			$this->_layout = PHPPowerPoint_DocumentLayout::LAYOUT_CUSTOM;
			break;
	}

    	return $this;
    }

    /**
     * Get Document Layout cx
     *
     * @return integer
     */
    public function getCX() {
    	return $this->_cx;
    }

    /**
     * Get Document Layout cy
     *
     * @return integer
     */
    public function getCY() {
    	return $this->_cy;
    }

    /**
     * Get Document Layout in millimeters
     *
     * @return integer
     */
    public function getLayoutXmilli() {
    	return $this->_cx / 36000;
    }

    /**
     * Get Document Layout in millimeters
     *
     * @return integer
     */
    public function getLayoutYmilli() {
    	return $this->_cy / 36000;
    }

    /**
     * Set Document Layout in millimeters
     *
     * @param integer $pValue	Layout width
     * @return PHPPowerPoint_DocumentLayout
     */
    public function setLayoutXmilli(integer $pValue) {
	$this->_cx = $pValue * 36000;
	$this->_layout = PHPPowerPoint_DocumentLayout::LAYOUT_CUSTOM;
    	return $this;
    }
    /**
     * Set Document Layout in millimeters
     *
     * @param integer $pValue	Layout height
     * @return PHPPowerPoint_DocumentLayout
     */
    public function setLayoutYmilli(integer $pValue) {
	$this->_cx = $pValue * 36000;
	$this->_layout = PHPPowerPoint_DocumentLayout::LAYOUT_CUSTOM;
    	return $this;
    }

	/**
	 * Implement PHP __clone to create a deep clone, not just a shallow copy.
	 */
	public function __clone() {
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
