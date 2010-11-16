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
 * PHPPowerPoint_Style_Alignment
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Style
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Style_Alignment implements PHPPowerPoint_IComparable
{
	/* Horizontal alignment styles */
	const HORIZONTAL_GENERAL				= 'l';
	const HORIZONTAL_LEFT					= 'l';
	const HORIZONTAL_RIGHT					= 'r';
	const HORIZONTAL_CENTER					= 'ctr';
	const HORIZONTAL_JUSTIFY				= 'just';
	const HORIZONTAL_DISTRIBUTED			= 'dist';

	/* Vertical alignment styles */
	const VERTICAL_BASE						= 'base';
	const VERTICAL_AUTO						= 'auto';
	const VERTICAL_BOTTOM					= 'b';
	const VERTICAL_TOP						= 't';
	const VERTICAL_CENTER					= 'ctr';

	/**
	 * Horizontal
	 *
	 * @var string
	 */
	private $_horizontal;

	/**
	 * Vertical
	 *
	 * @var string
	 */
	private $_vertical;

	/**
	 * Level
	 *
	 * @var int
	 */
	private $_level;

	/**
	 * Indent - only possible with horizontal alignment left and right
	 *
	 * @var int
	 */
	private $_indent;

	/**
	 * Margin left - only possible with horizontal alignment left and right
	 *
	 * @var int
	 */
	private $_marginLeft;

	/**
	 * Margin right - only possible with horizontal alignment left and right
	 *
	 * @var int
	 */
	private $_marginRight;

    /**
     * Create a new PHPPowerPoint_Style_Alignment
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_horizontal			= PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT;
    	$this->_vertical			= PHPPowerPoint_Style_Alignment::VERTICAL_BASE;
    	$this->_level				= 0;
		$this->_indent				= 0;
		$this->_marginLeft			= 0;
		$this->_marginRight			= 0;
    }

    /**
     * Get Horizontal
     *
     * @return string
     */
    public function getHorizontal() {
    	return $this->_horizontal;
    }

    /**
     * Set Horizontal
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Alignment
     */
    public function setHorizontal($pValue = PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT) {
        if ($pValue == '') {
    		$pValue = PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT;
    	}
    	$this->_horizontal = $pValue;
    	return $this;
    }

    /**
     * Get Vertical
     *
     * @return string
     */
    public function getVertical() {
    	return $this->_vertical;
    }

    /**
     * Set Vertical
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Alignment
     */
    public function setVertical($pValue = PHPPowerPoint_Style_Alignment::VERTICAL_BASE) {
    	if ($pValue == '') {
    		$pValue = PHPPowerPoint_Style_Alignment::VERTICAL_BASE;
    	}
    	$this->_vertical = $pValue;
    	return $this;
    }

    /**
     * Get Level
     *
     * @return int
     */
    public function getLevel() {
    	return $this->_level;
    }

    /**
     * Set Level
     *
     * @param int $pValue Ranging 0 - 8
     * @throws Exception
     * @return PHPPowerPoint_Style_Alignment
     */
    public function setLevel($pValue = 0) {
    	if ($pValue < 0 || $pValue > 8) {
    		throw new Exception("Invalid value: shoul be range 0 - 8.");
    	}
    	$this->_level = $pValue;
    	return $this;
    }

    /**
     * Get indent
     *
     * @return int
     */
    public function getIndent() {
    	return $this->_indent;
    }

    /**
     * Set indent
     *
     * @param int $pValue
     * @return PHPPowerPoint_Style_Alignment
     */
    public function setIndent($pValue = 0) {
		if ($pValue > 0) {
			if ($this->getHorizontal() != self::HORIZONTAL_GENERAL && $this->getHorizontal() != self::HORIZONTAL_LEFT && $this->getHorizontal() != self::HORIZONTAL_RIGHT) {
				$pValue = 0; // indent not supported
			}
		}

		$this->_indent = $pValue;
		return $this;
    }

    /**
     * Get margin left
     *
     * @return int
     */
    public function getMarginLeft() {
    	return $this->_marginLeft;
    }

    /**
     * Set margin left
     *
     * @param int $pValue
     * @return PHPPowerPoint_Style_Alignment
     */
    public function setMarginLeft($pValue = 0) {
		if ($pValue > 0) {
			if ($this->getHorizontal() != self::HORIZONTAL_GENERAL && $this->getHorizontal() != self::HORIZONTAL_LEFT && $this->getHorizontal() != self::HORIZONTAL_RIGHT) {
				$pValue = 0; // margin left not supported
			}
		}

		$this->_marginLeft = $pValue;
		return $this;
    }

    /**
     * Get margin right
     *
     * @return int
     */
    public function getMarginRight() {
    	return $this->_marginRight;
    }

    /**
     * Set margin ight
     *
     * @param int $pValue
     * @return PHPPowerPoint_Style_Alignment
     */
    public function setMarginRight($pValue = 0) {
		if ($pValue > 0) {
			if ($this->getHorizontal() != self::HORIZONTAL_GENERAL && $this->getHorizontal() != self::HORIZONTAL_LEFT && $this->getHorizontal() != self::HORIZONTAL_RIGHT) {
				$pValue = 0; // margin left not supported
			}
		}

		$this->_marginRight = $pValue;
		return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_horizontal
    		. $this->_vertical
    		. $this->_level
			. $this->_indent
			. $this->_marginLeft
			. $this->_marginRight
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
