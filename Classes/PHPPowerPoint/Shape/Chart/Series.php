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
 * PHPPowerPoint_Shape_Chart_Series
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape_Chart
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Chart_Series implements PHPPowerPoint_IComparable
{
	/**
	 * Title
	 *
	 * @var string
	 */
	private $_title = 'Series Title';
	
	/**
	 * Fill
	 *
	 * @var PHPPowerPoint_Style_Fill
	 */
	private $_fill;
	
	/**
	 * Values (key/value)
	 *
	 * @var array
	 */
	private $_values = array();
    
    /**
     * Create a new PHPPowerPoint_Shape_Chart_Series instance
     * 
     * @param string $title  Title
     * @param array  $values Values
     */
    public function __construct($title = 'Series Title', $values = array())
    {
    	$this->_fill   = new PHPPowerPoint_Style_Fill();
    	$this->_title  = $title;
    	$this->_values = $values;
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
	 * @return PHPPowerPoint_Shape_Chart_Series
	 */
	public function setTitle($value = 'Series Title') {
	        $this->_title = $value;
	        return $this;
	}
	
    /**
     * Get Fill
     *
     * @return PHPPowerPoint_Style_Fill
     */
    public function getFill() {
		return $this->_fill;
    }
    
	/**
	 * Get Values
	 *
	 * @return array
	 */
	public function getValues() {
	        return $this->_values;
	}
	
	/**
	 * Set Values
	 *
	 * @param array $value
	 * @return PHPPowerPoint_Shape_Chart_Series
	 */
	public function setValues($value = array()) {
	        $this->_values = $value;
	        return $this;
	}
	
	/**
	 * Add Value
	 *
	 * @param mixed $key
	 * @param mixed $value
	 * @return PHPPowerPoint_Shape_Chart_Series
	 */
	public function addValue($key, $value) {
	        $this->_values[$key] = $value;
	        return $this;
	}

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  (is_null($this->_fill) ? 'null' : $this->_fill->getHashCode())
    		. $this->_title
    		. var_export($this->_values, true)
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