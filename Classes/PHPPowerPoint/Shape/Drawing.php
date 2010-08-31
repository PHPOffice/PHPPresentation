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
 * PHPPowerPoint_Shape_Drawing
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Drawing extends PHPPowerPoint_Shape_BaseDrawing implements PHPPowerPoint_IComparable
{
	/**
	 * Path
	 *
	 * @var string
	 */
	private $_path;

    /**
     * Create a new PHPPowerPoint_Slide_Drawing
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_path				= '';

    	// Initialize parent
    	parent::__construct();
    }

    /**
     * Get Filename
     *
     * @return string
     */
    public function getFilename() {
    	return basename($this->_path);
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename() {
    	return str_replace('.' . $this->getExtension(), '', $this->getFilename()) . $this->getImageIndex() . '.' . $this->getExtension();
    }

    /**
     * Get Extension
     *
     * @return string
     */
    public function getExtension() {
    	$exploded = explode(".", basename($this->_path));
    	return $exploded[count($exploded) - 1];
    }

    /**
     * Get Path
     *
     * @return string
     */
    public function getPath() {
    	return $this->_path;
    }

    /**
     * Set Path
     *
     * @param 	string 		$pValue			File path
     * @param 	boolean		$pVerifyFile	Verify file
     * @throws 	Exception
     * @return PHPPowerPoint_Shape_Drawing
     */
    public function setPath($pValue = '', $pVerifyFile = true) {
    	if ($pVerifyFile) {
	    	if (file_exists($pValue)) {
	    		$this->_path = $pValue;

	    		if ($this->_width == 0 && $this->_height == 0) {
	    			// Get width/height
	    			list($this->_width, $this->_height) = getimagesize($pValue);
	    		}
	    	} else {
	    		throw new Exception("File $pValue not found!");
	    	}
    	} else {
    		$this->_path = $pValue;
    	}
    	return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_path
    		. parent::getHashCode()
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
