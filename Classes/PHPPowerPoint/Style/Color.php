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
 * PHPPowerPoint_Style_Color
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Style
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Style_Color implements PHPPowerPoint_IComparable
{
	/* Colors */
	const COLOR_BLACK						= 'FF000000';
	const COLOR_WHITE						= 'FFFFFFFF';
	const COLOR_RED							= 'FFFF0000';
	const COLOR_DARKRED						= 'FF800000';
	const COLOR_BLUE						= 'FF0000FF';
	const COLOR_DARKBLUE					= 'FF000080';
	const COLOR_GREEN						= 'FF00FF00';
	const COLOR_DARKGREEN					= 'FF008000';
	const COLOR_YELLOW						= 'FFFFFF00';
	const COLOR_DARKYELLOW					= 'FF808000';

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
    public function __construct($pARGB = PHPPowerPoint_Style_Color::COLOR_BLACK)
    {
    	// Initialise values
    	$this->_argb			= $pARGB;
    }

    /**
     * Get ARGB
     *
     * @return string
     */
    public function getARGB() {
    	return $this->_argb;
    }

    /**
     * Set ARGB
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Color
     */
    public function setARGB($pValue = PHPPowerPoint_Style_Color::COLOR_BLACK) {
    	if ($pValue == '') {
    		$pValue = PHPPowerPoint_Style_Color::COLOR_BLACK;
    	}
    	$this->_argb = $pValue;
    	return $this;
    }

    /**
     * Get RGB
     *
     * @return string
     */
    public function getRGB() {
    	if (strlen($this->_argb) == 6) {
    		return $this->_argb;
    	} else {
    		return substr($this->_argb, 2);
    	}
    }

    /**
     * Set RGB
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Color
     */
    public function setRGB($pValue = '000000') {
        if ($pValue == '') {
    		$pValue = '000000';
    	}
    	$this->_argb = 'FF' . $pValue;
    	return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
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
