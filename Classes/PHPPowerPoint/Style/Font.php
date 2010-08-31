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
 * PHPPowerPoint_Style_Font
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Style
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Style_Font implements PHPPowerPoint_IComparable
{
	/* Underline types */
	const UNDERLINE_NONE					= 'none';
	const UNDERLINE_DASH					= 'dash';
	const UNDERLINE_DASHHEAVY				= 'dashHeavy';
	const UNDERLINE_DASHLONG				= 'dashLong';
	const UNDERLINE_DASHLONGHEAVY			= 'dashLongHeavy';
	const UNDERLINE_DOUBLE					= 'dbl';
	const UNDERLINE_DOTHASH					= 'dotDash';
	const UNDERLINE_DOTHASHHEAVY			= 'dotDashHeavy';
	const UNDERLINE_DOTDOTDASH				= 'dotDotDash';
	const UNDERLINE_DOTDOTDASHHEAVY			= 'dotDotDashHeavy';
	const UNDERLINE_DOTTED					= 'dotted';
	const UNDERLINE_DOTTEDHEAVY				= 'dottedHeavy';
	const UNDERLINE_HEAVY					= 'heavy';
	const UNDERLINE_SINGLE					= 'sng';
	const UNDERLINE_WAVY					= 'wavy';
	const UNDERLINE_WAVYDOUBLE				= 'wavyDbl';
	const UNDERLINE_WAVYHEAVY				= 'wavyHeavy';
	const UNDERLINE_WORDS					= 'words';

	/**
	 * Name
	 *
	 * @var string
	 */
	private $_name;

	/**
	 * Bold
	 *
	 * @var boolean
	 */
	private $_bold;

	/**
	 * Italic
	 *
	 * @var boolean
	 */
	private $_italic;

	/**
	 * Superscript
	 *
	 * @var boolean
	 */
	private $_superScript;

	/**
	 * Subscript
	 *
	 * @var boolean
	 */
	private $_subScript;

	/**
	 * Underline
	 *
	 * @var string
	 */
	private $_underline;

	/**
	 * Strikethrough
	 *
	 * @var boolean
	 */
	private $_strikethrough;

	/**
	 * Foreground color
	 *
	 * @var PHPPowerPoint_Style_Color
	 */
	private $_color;

	/**
     * Create a new PHPPowerPoint_Style_Font
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_name				= 'Calibri';
    	$this->_size				= 10;
		$this->_bold				= false;
		$this->_italic				= false;
		$this->_superScript			= false;
		$this->_subScript			= false;
		$this->_underline			= PHPPowerPoint_Style_Font::UNDERLINE_NONE;
		$this->_strikethrough		= false;
		$this->_color				= new PHPPowerPoint_Style_Color(PHPPowerPoint_Style_Color::COLOR_BLACK);
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName() {
    	return $this->_name;
    }

    /**
     * Set Name
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Font
     */
    public function setName($pValue = 'Calibri') {
   		if ($pValue == '') {
    		$pValue = 'Calibri';
    	}
    	$this->_name = $pValue;
    	return $this;
    }

    /**
     * Get Size
     *
     * @return double
     */
    public function getSize() {
    	return $this->_size;
    }

    /**
     * Set Size
     *
     * @param double $pValue
     * @return PHPPowerPoint_Style_Font
     */
    public function setSize($pValue = 10) {
    	if ($pValue == '') {
    		$pValue = 10;
    	}
    	$this->_size = $pValue;
    	return $this;
    }

    /**
     * Get Bold
     *
     * @return boolean
     */
    public function getBold() {
    	return $this->_bold;
    }

    /**
     * Set Bold
     *
     * @param boolean $pValue
     * @return PHPPowerPoint_Style_Font
     */
    public function setBold($pValue = false) {
    	if ($pValue == '') {
    		$pValue = false;
    	}
    	$this->_bold = $pValue;
    	return $this;
    }

    /**
     * Get Italic
     *
     * @return boolean
     */
    public function getItalic() {
    	return $this->_italic;
    }

    /**
     * Set Italic
     *
     * @param boolean $pValue
     * @return PHPPowerPoint_Style_Font
     */
    public function setItalic($pValue = false) {
    	if ($pValue == '') {
    		$pValue = false;
    	}
    	$this->_italic = $pValue;
    	return $this;
    }

    /**
     * Get SuperScript
     *
     * @return boolean
     */
    public function getSuperScript() {
    	return $this->_superScript;
    }

    /**
     * Set SuperScript
     *
     * @param boolean $pValue
     * @return PHPPowerPoint_Style_Font
     */
    public function setSuperScript($pValue = false) {
    	if ($pValue == '') {
    		$pValue = false;
    	}
    	$this->_superScript = $pValue;
		$this->_subScript = !$pValue;
		return $this;
    }

	/**
     * Get SubScript
     *
     * @return boolean
     */
    public function getSubScript() {
    	return $this->_subScript;
    }

    /**
     * Set SubScript
     *
     * @param boolean $pValue
     * @return PHPPowerPoint_Style_Font
     */
    public function setSubScript($pValue = false) {
    	if ($pValue == '') {
    		$pValue = false;
    	}
    	$this->_subScript = $pValue;
		$this->_superScript = !$pValue;
		return $this;
    }

    /**
     * Get Underline
     *
     * @return string
     */
    public function getUnderline() {
    	return $this->_underline;
    }

    /**
     * Set Underline
     *
     * @param string $pValue	PHPPowerPoint_Style_Font underline type
     * @return PHPPowerPoint_Style_Font
     */
    public function setUnderline($pValue = PHPPowerPoint_Style_Font::UNDERLINE_NONE) {
    	if ($pValue == '') {
    		$pValue = PHPPowerPoint_Style_Font::UNDERLINE_NONE;
    	}
    	$this->_underline = $pValue;
    	return $this;
    }

    /**
     * Get Striketrough
     *
     * @deprecated Use getStrikethrough() instead.
     * @return boolean
     */
    public function getStriketrough() {
    	return $this->getStrikethrough();
    }

    /**
     * Set Striketrough
     *
     * @deprecated Use setStrikethrough() instead.
     * @param boolean $pValue
     * @return PHPPowerPoint_Style_Font
     */
    public function setStriketrough($pValue = false) {
    	return $this->setStrikethrough($pValue);
    }

    /**
     * Get Strikethrough
     *
     * @return boolean
     */
    public function getStrikethrough() {
    	return $this->_strikethrough;
    }

    /**
     * Set Strikethrough
     *
     * @param boolean $pValue
     * @return PHPPowerPoint_Style_Font
     */
    public function setStrikethrough($pValue = false) {
    	if ($pValue == '') {
    		$pValue = false;
    	}
    	$this->_strikethrough = $pValue;
    	return $this;
    }

    /**
     * Get Color
     *
     * @return PHPPowerPoint_Style_Color
     */
    public function getColor() {
    	return $this->_color;
    }

    /**
     * Set Color
     *
     * @param 	PHPPowerPoint_Style_Color $pValue
     * @throws 	Exception
     * @return PHPPowerPoint_Style_Font
     */
    public function setColor(PHPPowerPoint_Style_Color $pValue = null) {
   		$this->_color = $pValue;
   		return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_name
    		. $this->_size
    		. ($this->_bold ? 't' : 'f')
    		. ($this->_italic ? 't' : 'f')
			. ($this->_superScript ? 't' : 'f')
			. ($this->_subScript ? 't' : 'f')
    		. $this->_underline
    		. ($this->_strikethrough ? 't' : 'f')
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
