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
 * PHPPowerPoint_Style_Border
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Style
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Style_Border implements PHPPowerPoint_IComparable
{
	/* Line style */
	const LINE_NONE				= 'none';
	const LINE_SINGLE  			= 'sng';
	const LINE_DOUBLE			= 'dbl';
	const LINE_THICKTHIN		= 'thickThin';
	const LINE_THINTHICK		= 'thinThick';
	const LINE_TRI				= 'tri';

	/* Dash style */
	const DASH_DASH				= 'dash';
	const DASH_DASHDOT			= 'dashDot';
	const DASH_DOT				= 'dot';
	const DASH_LARGEDASH		= 'lgDash';
	const DASH_LARGEDASHDOT		= 'lgDashDot';
	const DASH_LARGEDASHDOTDOT	= 'lgDashDotDot';
	const DASH_SOLID			= 'solid';
	const DASH_SYSDASH			= 'sysDash';
	const DASH_SYSDASHDOT		= 'sysDashDot';
	const DASH_SYSDASHDOTDOT	= 'sysDashDotDot';
	const DASH_SYSDOT			= 'sysDot';
	
	/**
	 * Line width
	 *
	 * @var int
	 */
	private $_lineWidth = 1;

	/**
	 * Line style
	 *
	 * @var string
	 */
	private $_lineStyle;

	/**
	 * Dash style
	 *
	 * @var string
	 */
	private $_dashStyle;

	/**
	 * Border color
	 *
	 * @var PHPPowerPoint_Style_Color
	 */
	private $_color;

    /**
     * Create a new PHPPowerPoint_Style_Border
     */
    public function __construct()
    {
    	// Initialise values
		$this->_lineWidth = 1;
		$this->_lineStyle = self::LINE_SINGLE;
		$this->_dashStyle = self::DASH_SOLID;
		$this->_color	  = new PHPPowerPoint_Style_Color(PHPPowerPoint_Style_Color::COLOR_BLACK);
    }

    /**
     * Get line width
     *
     * @return int
     */
    public function getLineWidth() {
    	return $this->_lineWidth;
    }

    /**
     * Set line width
     *
     * @param int $pValue
     * @return PHPPowerPoint_Style_Border
     */
    public function setLineWidth($pValue = 1) {
		$this->_lineWidth = $pValue;
		return $this;
    }

    /**
     * Get line style
     *
     * @return string
     */
    public function getLineStyle() {
    	return $this->_lineStyle;
    }

    /**
     * Set line style
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Border
     */
    public function setLineStyle($pValue = PHPPowerPoint_Style_Border::LINE_SINGLE) {
        if ($pValue == '') {
    		$pValue = PHPPowerPoint_Style_Border::LINE_SINGLE;
    	}
		$this->_lineStyle = $pValue;
		return $this;
    }

    /**
     * Get dash style
     *
     * @return string
     */
    public function getDashStyle() {
    	return $this->_dashStyle;
    }

    /**
     * Set dash style
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Border
     */
    public function setDashStyle($pValue = PHPPowerPoint_Style_Border::DASH_SOLID) {
        if ($pValue == '') {
    		$pValue = PHPPowerPoint_Style_Border::DASH_SOLID;
    	}
		$this->_dashStyle = $pValue;
		return $this;
    }

    /**
     * Get Border Color
     *
     * @return PHPPowerPoint_Style_Color
     */
    public function getColor() {
    	return $this->_color;
    }

    /**
     * Set Border Color
     *
     * @param 	PHPPowerPoint_Style_Color $color
     * @throws 	Exception
     * @return PHPPowerPoint_Style_Border
     */
    public function setColor(PHPPowerPoint_Style_Color $color = null) {
		$this->_color = $color;
		return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_lineStyle
    		. $this->_lineWidth
    		. $this->_dashStyle
    		. $this->_color->getHashCode()
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
