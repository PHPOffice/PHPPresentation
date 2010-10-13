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
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Shape_Chart
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Chart extends PHPPowerPoint_Shape_BaseDrawing implements PHPPowerPoint_IComparable
{
	/**
	 * Title
	 *
	 * @var PHPPowerPoint_Shape_Chart_Title
	 */
	private $_title;
	
	/**
	 * Legend
	 *
	 * @var PHPPowerPoint_Shape_Chart_Legend
	 */
	private $_legend;
	
	/**
	 * Plot area
	 *
	 * @var PHPPowerPoint_Shape_Chart_PlotArea
	 */
	private $_plotArea;
	
	/**
	 * View 3D
	 *
	 * @var PHPPowerPoint_Shape_Chart_View3D
	 */
	private $_view3D;

    /**
     * Create a new PHPPowerPoint_Slide_MemoryDrawing
     */
    public function __construct()
    {
    	// Initialize
    	$this->_title    = new PHPPowerPoint_Shape_Chart_Title();
    	$this->_legend   = new PHPPowerPoint_Shape_Chart_Legend();
    	$this->_plotArea = new PHPPowerPoint_Shape_Chart_PlotArea();
    	$this->_view3D   = new PHPPowerPoint_Shape_Chart_View3D();
    	
    	// Initialize parent
    	parent::__construct();
    }
    
	/**
	 * Get Title
	 *
	 * @return PHPPowerPoint_Shape_Chart_Title
	 */
	public function getTitle() {
	        return $this->_title;
	}
	
	/**
	 * Get Legend
	 *
	 * @return PHPPowerPoint_Shape_Chart_Legend
	 */
	public function getLegend() {
	        return $this->_legend;
	}
	
	/**
	 * Get PlotArea
	 *
	 * @return PHPPowerPoint_Shape_Chart_PlotArea
	 */
	public function getPlotArea() {
	        return $this->_plotArea;
	}
	
	/**
	 * Get View3D
	 *
	 * @return PHPPowerPoint_Shape_Chart_View3D
	 */
	public function getView3D() {
	        return $this->_view3D;
	}

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename() {
    	return 'chart' . $this->getImageIndex() . '.xml';
    }
    
	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  parent::getHashCode()
    		. $this->_title->getHashCode()
    		. $this->_legend->getHashCode()
    		. $this->_plotArea->getHashCode()
    		. $this->_view3D->getHashCode()
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
