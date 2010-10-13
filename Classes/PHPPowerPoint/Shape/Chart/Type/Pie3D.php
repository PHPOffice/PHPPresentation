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
 * PHPPowerPoint_Shape_Chart_Type_Pie3D
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape_Chart_Type
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Chart_Type_Pie3D extends PHPPowerPoint_Shape_Chart_Type implements PHPPowerPoint_IComparable
{
	/**
	 * Data
	 *
	 * @var array
	 */
	private $_data = array();
    
    /**
     * Create a new PHPPowerPoint_Shape_Chart_Type_Pie3D instance
     */
    public function __construct()
    {
    	$this->_hasAxisX = false;
    	$this->_hasAxisY = false;
    }

	/**
	 * Get Data
	 *
	 * @return array
	 */
	public function getData() {
	        return $this->_data;
	}
	
	/**
	 * Set Data
	 *
	 * @param array $value Array of PHPPowerPoint_Shape_Chart_Series
	 * @return PHPPowerPoint_Shape_Chart_Type_Pie3D
	 */
	public function setData($value = array()) {
	        $this->_data = $value;
	        return $this;
	}
	
	/**
	 * Add Series
	 *
	 * @param PHPPowerPoint_Shape_Chart_Series $value
	 * @return PHPPowerPoint_Shape_Chart_Type_Pie3D
	 */
	public function addSeries(PHPPowerPoint_Shape_Chart_Series $value) {
	        $this->_data[] = $value;
	        return $this;
	}
	
	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
		$hash = '';
		foreach ($this->_data as $series) {
			$hash .= $series->getHashCode();
		}
		
    	return md5(
    		  $hash
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