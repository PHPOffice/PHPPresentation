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
 * @package    PHPPowerPoint_Shape_Chart
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Shape_Chart_Axis
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape_Chart
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Chart_Axis implements PHPPowerPoint_IComparable
{	
	/**
	 * Title
	 *
	 * @var string
	 */
	private $_title = 'Axis Title';
	
	/**
	 * Format code
	 *
	 * @var string
	 */
	private $_formatCode = '';

    /**
     * Create a new PHPPowerPoint_Shape_Chart_Axis instance
     * 
     * @param string $title Title
     */
    public function __construct($title = 'Axis Title')
    {
    	$this->_title = $title;
    }
    
	/**
	 * Get Title
	 *
	 * @return string
	 */
	public function getTitle() {
	        return $this->_title;
	}
	
	/**
	 * Set Title
	 *
	 * @param string $value
	 * @return PHPPowerPoint_Shape_Chart_Axis
	 */
	public function setTitle($value = 'Axis Title') {
	        $this->_title = $value;
	        return $this;
	}
	
	/**
	 * Get Format Code
	 *
	 * @return string
	 */
	public function getFormatCode() {
	        return $this->_formatCode;
	}
	
	/**
	 * Set Format Code
	 *
	 * @param string $value
	 * @return PHPPowerPoint_Shape_Chart_Axis
	 */
	public function setFormatCode($value = '') {
	        $this->_formatCode = $value;
	        return $this;
	}
	
	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_title
    		. $this->_formatCode
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