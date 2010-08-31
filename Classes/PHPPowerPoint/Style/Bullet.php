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
 * PHPPowerPoint_Style_Bullet
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Style
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Style_Bullet implements PHPPowerPoint_IComparable
{
	/* Bullet types */
	const TYPE_NONE							= 'none';
	const TYPE_BULLET						= 'bullet';
	const TYPE_NUMERIC						= 'numeric';

	/* Numeric bullet styles */
	const NUMERIC_DEFAULT					= 'arabicPeriod';
	const NUMERIC_ALPHALCPARENBOTH 			= 'alphaLcParenBoth';
	const NUMERIC_ALPHAUCPARENBOTH			= 'alphaUcParenBoth';
	const NUMERIC_ALPHALCPARENR 			= 'alphaLcParenR';
	const NUMERIC_ALPHAUCPARENR 			= 'alphaUcParenR';
	const NUMERIC_ALPHALCPERIOD 			= 'alphaLcPeriod';
	const NUMERIC_ALPHAUCPERIOD 			= 'alphaUcPeriod';
	const NUMERIC_ARABICPARENBOTH 			= 'arabicParenBoth';
	const NUMERIC_ARABICPARENR 				= 'arabicParenR';
	const NUMERIC_ARABICPERIOD 				= 'arabicPeriod';
	const NUMERIC_ARABICPLAIN 				= 'arabicPlain';
	const NUMERIC_ROMANLCPARENBOTH			= 'romanLcParenBoth';
	const NUMERIC_ROMANUCPARENBOTH 			= 'romanUcParenBoth';
	const NUMERIC_ROMANLCPARENR 			= 'romanLcParenR';
	const NUMERIC_ROMANUCPARENR 			= 'romanUcParenR';
	const NUMERIC_ROMANLCPERIOD 			= 'romanLcPeriod';
	const NUMERIC_ROMANUCPERIOD 			= 'romanUcPeriod';
	const NUMERIC_CIRCLENUMDBPLAIN 			= 'circleNumDbPlain';
	const NUMERIC_CIRCLENUMWDBLACKPLAIN 	= 'circleNumWdBlackPlain';
	const NUMERIC_CIRCLENUMWDWHITEPLAIN 	= 'circleNumWdWhitePlain';
	const NUMERIC_ARABICDBPERIOD 			= 'arabicDbPeriod';
	const NUMERIC_ARABICDBPLAIN 			= 'arabicDbPlain';
	const NUMERIC_EA1CHSPERIOD 				= 'ea1ChsPeriod';
	const NUMERIC_EA1CHSPLAIN 				= 'ea1ChsPlain';
	const NUMERIC_EA1CHTPERIOD 				= 'ea1ChtPeriod';
	const NUMERIC_EA1CHTPLAIN 				= 'ea1ChtPlain';
	const NUMERIC_EA1JPNCHSDBPERIOD 		= 'ea1JpnChsDbPeriod';
	const NUMERIC_EA1JPNKORPLAIN 			= 'ea1JpnKorPlain';
	const NUMERIC_EA1JPNKORPERIOD 			= 'ea1JpnKorPeriod';
	const NUMERIC_ARABIC1MINUS 				= 'arabic1Minus';
	const NUMERIC_ARABIC2MINUS 				= 'arabic2Minus';
	const NUMERIC_HEBREW2MINUS 				= 'hebrew2Minus';
	const NUMERIC_THAIALPHAPERIOD 			= 'thaiAlphaPeriod';
	const NUMERIC_THAIALPHAPARENR 			= 'thaiAlphaParenR';
	const NUMERIC_THAIALPHAPARENBOTH 		= 'thaiAlphaParenBoth';
	const NUMERIC_THAINUMPERIOD 			= 'thaiNumPeriod';
	const NUMERIC_THAINUMPARENR 			= 'thaiNumParenR';
	const NUMERIC_THAINUMPARENBOTH			= 'thaiNumParenBoth';
	const NUMERIC_HINDIALPHAPERIOD			= 'hindiAlphaPeriod';
	const NUMERIC_HINDINUMPERIOD 			= 'hindiNumPeriod';
	const NUMERIC_HINDINUMPARENR 			= 'hindiNumParenR';
	const NUMERIC_HINDIALPHA1PERIOD 		= 'hindiAlpha1Period';

	/**
	 * Bullet type
	 *
	 * @var string
	 */
	private $_bulletType = self::TYPE_NONE;

	/**
	 * Bullet font
	 *
	 * @var string
	 */
	private $_bulletFont;

	/**
	 * Bullet char
	 *
	 * @var string
	 */
	private $_bulletChar = '-';

	/**
	 * Bullet numeric style
	 *
	 * @var string
	 */
	private $_bulletNumericStyle = self::NUMERIC_DEFAULT;

	/**
	 * Bullet numeric start at
	 *
	 * @var int
	 */
	private $_bulletNumericStartAt = 1;

	/**
     * Create a new PHPPowerPoint_Style_Bullet
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_bulletType				= self::TYPE_NONE;
    	$this->_bulletFont				= 'Calibri';
    	$this->_bulletChar				= '-';
    	$this->_bulletNumericStyle		= self::NUMERIC_DEFAULT;
    	$this->_bulletNumericStartAt 	= 1;
    }

    /**
     * Get bullet type
     *
     * @return string
     */
    public function getBulletType() {
    	return $this->_bulletType;
    }

    /**
     * Set bullet type
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Bullet
     */
    public function setBulletType($pValue = self::TYPE_NONE) {
    	$this->_bulletType = $pValue;
    	return $this;
    }

    /**
     * Get bullet font
     *
     * @return string
     */
    public function getBulletFont() {
    	return $this->_bulletFont;
    }

    /**
     * Set bullet font
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Bullet
     */
    public function setBulletFont($pValue = 'Calibri') {
   		if ($pValue == '') {
    		$pValue = 'Calibri';
    	}
    	$this->_bulletFont = $pValue;
    	return $this;
    }

    /**
     * Get bullet char
     *
     * @return string
     */
    public function getBulletChar() {
    	return $this->_bulletChar;
    }

    /**
     * Set bullet char
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Bullet
     */
    public function setBulletChar($pValue = '-') {
    	$this->_bulletChar = $pValue;
    	return $this;
    }

    /**
     * Get bullet numeric style
     *
     * @return string
     */
    public function getBulletNumericStyle() {
    	return $this->_bulletNumericStyle;
    }

    /**
     * Set bullet numeric style
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Bullet
     */
    public function setBulletNumericStyle($pValue = self::NUMERIC_DEFAULT) {
    	$this->_bulletNumericStyle = $pValue;
    	return $this;
    }

    /**
     * Get bullet numeric start at
     *
     * @return string
     */
    public function getBulletNumericStartAt() {
    	return $this->_bulletNumericStartAt;
    }

    /**
     * Set bullet numeric start at
     *
     * @param string $pValue
     * @return PHPPowerPoint_Style_Bullet
     */
    public function setBulletNumericStartAt($pValue = 1) {
    	$this->_bulletNumericStartAt = $pValue;
    	return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_bulletType
    		. $this->_bulletFont
    		. $this->_bulletChar
    		. $this->_bulletNumericStyle
    		. $this->_bulletNumericStartAt
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
