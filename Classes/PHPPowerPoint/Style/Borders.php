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
 * PHPPowerPoint_Style_Borders
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Style
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Style_Borders implements PHPPowerPoint_IComparable
{
	/**
	 * Left
	 *
	 * @var PHPPowerPoint_Style_Border
	 */
	private $_left;

	/**
	 * Right
	 *
	 * @var PHPPowerPoint_Style_Border
	 */
	private $_right;

	/**
	 * Top
	 *
	 * @var PHPPowerPoint_Style_Border
	 */
	private $_top;

	/**
	 * Bottom
	 *
	 * @var PHPPowerPoint_Style_Border
	 */
	private $_bottom;

	/**
	 * Diagonal up
	 *
	 * @var PHPPowerPoint_Style_Border
	 */
	private $_diagonalUp;

	/**
	 * Diagonal down
	 *
	 * @var PHPPowerPoint_Style_Border
	 */
	private $_diagonalDown;

	/**
     * Create a new PHPPowerPoint_Style_Borders
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_left				= new PHPPowerPoint_Style_Border();
    	$this->_right				= new PHPPowerPoint_Style_Border();
    	$this->_top					= new PHPPowerPoint_Style_Border();
    	$this->_bottom				= new PHPPowerPoint_Style_Border();
    	$this->_diagonalUp			= new PHPPowerPoint_Style_Border();
    	$this->_diagonalUp->setLineStyle(PHPPowerPoint_Style_Border::LINE_NONE);
    	$this->_diagonalDown		= new PHPPowerPoint_Style_Border();
    	$this->_diagonalDown->setLineStyle(PHPPowerPoint_Style_Border::LINE_NONE);
    }

    /**
     * Get Left
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getLeft() {
		return $this->_left;
    }

    /**
     * Get Right
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getRight() {
		return $this->_right;
    }

    /**
     * Get Top
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getTop() {
		return $this->_top;
    }

    /**
     * Get Bottom
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getBottom() {
		return $this->_bottom;
    }

    /**
     * Get Diagonal Up
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getDiagonalUp() {
		return $this->_diagonalUp;
    }

    /**
     * Get Diagonal Down
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getDiagonalDown() {
		return $this->_diagonalDown;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->getLeft()->getHashCode()
    		. $this->getRight()->getHashCode()
    		. $this->getTop()->getHashCode()
    		. $this->getBottom()->getHashCode()
    		. $this->getDiagonalUp()->getHashCode()
    		. $this->getDiagonalDown()->getHashCode()
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
