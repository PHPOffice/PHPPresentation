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
 * PHPPowerPoint_Shape_Table
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Table extends PHPPowerPoint_Shape_BaseDrawing implements PHPPowerPoint_IComparable
{
	/**
	 * Rows
	 *
	 * @var PHPPowerPoint_Shape_Table_Row[]
	 */
	private $_rows;

	/**
	 * Number of columns
	 *
	 * @var int
	 */
	private $_columnCount = 1;

    /**
     * Create a new PHPPowerPoint_Shape_Table instance
     *
     * @param int $columns Number of columns
     */
    public function __construct($columns = 1)
    {
    	// Initialise variables
    	$this->_rows = array();
    	$this->_columnCount = $columns;

    	// Initialize parent
    	parent::__construct();

    	// No resize proportional
    	$this->_resizeProportional	= false;
    }

    /**
     * Get row
     *
     * @param int $row Row number
     * @param boolean $exceptionAsNull Return a null value instead of an exception?
     * @return PHPPowerPoint_Shape_Table_Row
     */
	public function getRow($row = 0, $exceptionAsNull = false)
    {
    	if (!isset($this->_rows[$row])) {
    	    if ($exceptionAsNull) {
    			return null;
    		}
    		throw new Exception('Row number out of bounds.');
    	}
    	
    	return $this->_rows[$row];
    }

    /**
     * Get rows
     *
     * @return PHPPowerPoint_Shape_Table_Row[]
     */
	public function getRows()
    {
    	return $this->_rows;
    }

    /**
     * Create row
     *
     * @return PHPPowerPoint_Shape_Table_Row
     */
	public function createRow()
    {
    	$row = new PHPPowerPoint_Shape_Table_Row($this->_columnCount);
    	$this->_rows[] = $row;
    	return $row;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
		$hashElements = '';
		foreach ($this->_rows as $row) {
			$hashElements .= $row->getHashCode();
		}

    	return md5(
    		  $hashElements
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
			if ($key == '_parent') continue;

			if (is_object($value)) {
				$this->$key = clone $value;
			} else {
				$this->$key = $value;
			}
		}
	}
}
