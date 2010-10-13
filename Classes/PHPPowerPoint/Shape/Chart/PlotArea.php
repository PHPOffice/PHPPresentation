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
 * PHPPowerPoint_Shape_Chart_PlotArea
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape_Chart
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Chart_PlotArea implements PHPPowerPoint_IComparable
{
	/**
	 * Type
	 *
	 * @var PHPPowerPoint_Shape_Chart_Type
	 */
	private $_type;
	
	/**
	 * Axis X
	 *
	 * @var PHPPowerPoint_Shape_Chart_Axis
	 */
	private $_axisX;
	
	/**
	 * Axis Y
	 *
	 * @var PHPPowerPoint_Shape_Chart_Axis
	 */
	private $_axisY;
	
	/**
	 * OffsetX (as a fraction of the chart)
	 *
	 * @var float
	 */
	private $_offsetX = 0;
	
	/**
	 * OffsetY (as a fraction of the chart)
	 *
	 * @var float
	 */
	private $_offsetY = 0;
	
	/**
	 * Width (as a fraction of the chart)
	 *
	 * @var float
	 */
	private $_width = 0;
	
	/**
	 * Height (as a fraction of the chart)
	 *
	 * @var float
	 */
	private $_height = 0;
    
    /**
     * Create a new PHPPowerPoint_Shape_Chart_PlotArea instance
     */
    public function __construct()
    {
    	$this->_type  = null;
    	$this->_axisX = new PHPPowerPoint_Shape_Chart_Axis();
    	$this->_axisY = new PHPPowerPoint_Shape_Chart_Axis();
    }

	/**
	 * Get type
	 *
	 * @return PHPPowerPoint_Shape_Chart_Type
	 * @throws Exception
	 */
	public function getType() {
		if (is_null($this->_type)) {
			throw new Exception('Chart type has not been set.');
		}
	    return $this->_type;
	}
	
	/**
	 * Set type
	 *
	 * @param PHPPowerPoint_Shape_Chart_Type $value
	 * @return PHPPowerPoint_Shape_Chart_PlotArea
	 * @throws Exception
	 */
	public function setType(PHPPowerPoint_Shape_Chart_Type $value) {
		$this->_type = $value;
		return $this;
	}
	
	/**
	 * Get Axis X
	 *
	 * @return PHPPowerPoint_Shape_Chart_Axis
	 */
	public function getAxisX() {
		return $this->_axisX;
	}
	
	/**
	 * Get Axis Y
	 *
	 * @return PHPPowerPoint_Shape_Chart_Axis
	 */
	public function getAxisY() {
		return $this->_axisY;
	}
	
	/**
	 * Get OffsetX (as a fraction of the chart)
	 *
	 * @return float
	 */
	public function getOffsetX() {
	        return $this->_offsetX;
	}
	
	/**
	 * Set OffsetX (as a fraction of the chart)
	 *
	 * @param float $value
	 * @return PHPPowerPoint_Shape_Chart_Title
	 */
	public function setOffsetX($value = 0) {
	        $this->_offsetX = $value;
	        return $this;
	}
	
	/**
	 * Get OffsetY (as a fraction of the chart)
	 *
	 * @return float
	 */
	public function getOffsetY() {
	        return $this->_offsetY;
	}
	
	/**
	 * Set OffsetY (as a fraction of the chart)
	 *
	 * @param float $value
	 * @return PHPPowerPoint_Shape_Chart_Title
	 */
	public function setOffsetY($value = 0) {
	        $this->_offsetY = $value;
	        return $this;
	}
	
	/**
	 * Get Width (as a fraction of the chart)
	 *
	 * @return float
	 */
	public function getWidth() {
	        return $this->_width;
	}
	
	/**
	 * Set Width (as a fraction of the chart)
	 *
	 * @param float $value
	 * @return PHPPowerPoint_Shape_Chart_Title
	 */
	public function setWidth($value = 0) {
	        $this->_width = $value;
	        return $this;
	}
	
	/**
	 * Get Height (as a fraction of the chart)
	 *
	 * @return float
	 */
	public function getHeight() {
	        return $this->_height;
	}
	
	/**
	 * Set Height (as a fraction of the chart)
	 *
	 * @param float $value
	 * @return PHPPowerPoint_Shape_Chart_Title
	 */
	public function setHeight($value = 0) {
	        $this->_height = $value;
	        return $this;
	}
	
	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  (is_null($this->_type) ? 'null' : $this->_type->getHashCode())
    		. $this->_axisX->getHashCode()
    		. $this->_axisY->getHashCode()
    		. $this->_offsetX
    		. $this->_offsetY
    		. $this->_width
    		. $this->_height
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