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
 * PHPPowerPoint_Shape_BaseDrawing
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
abstract class PHPPowerPoint_Shape_BaseDrawing extends PHPPowerPoint_Shape implements PHPPowerPoint_IComparable
{
	/**
	 * Image counter
	 *
	 * @var int
	 */
	private static $_imageCounter = 0;

	/**
	 * Image index
	 *
	 * @var int
	 */
	private $_imageIndex = 0;

	/**
	 * Name
	 *
	 * @var string
	 */
	protected $_name;

	/**
	 * Description
	 *
	 * @var string
	 */
	protected $_description;

	/**
	 * Proportional resize
	 *
	 * @var boolean
	 */
	protected $_resizeProportional;
	
    /**
     * Slide relation ID (should not be used by user code!)
     * 
     * @var string
     */
    public $__relationId = null;

    /**
     * Create a new PHPPowerPoint_Slide_BaseDrawing
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_name				= '';
    	$this->_description			= '';
    	$this->_resizeProportional	= true;

		// Set image index
		self::$_imageCounter++;
		$this->_imageIndex 			= self::$_imageCounter;

    	// Initialize parent
    	parent::__construct();
    }

    /**
     * Get image index
     *
     * @return int
     */
    public function getImageIndex() {
    	return $this->_imageIndex;
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
     * @return PHPPowerPoint_Shape_BaseDrawing
     */
    public function setName($pValue = '') {
    	$this->_name = $pValue;
    	return $this;
    }

    /**
     * Get Description
     *
     * @return string
     */
    public function getDescription() {
    	return $this->_description;
    }

    /**
     * Set Description
     *
     * @param string $pValue
     * @return PHPPowerPoint_Shape_BaseDrawing
     */
    public function setDescription($pValue = '') {
    	$this->_description = $pValue;
    	return $this;
    }

    /**
     * Set Width
     *
     * @param int $pValue
     * @return PHPPowerPoint_Shape_BaseDrawing
     */
    public function setWidth($pValue = 0) {
    	// Resize proportional?
    	if ($this->_resizeProportional && $pValue != 0) {
    		$ratio = $this->_height / $this->_width;
    		$this->_height = round($ratio * $pValue);
    	}

    	// Set width
    	$this->_width = $pValue;

    	return $this;
    }

    /**
     * Set Height
     *
     * @param int $pValue
     * @return PHPPowerPoint_Shape_BaseDrawing
     */
    public function setHeight($pValue = 0) {
    	// Resize proportional?
    	if ($this->_resizeProportional && $pValue != 0) {
    		$ratio = $this->_width / $this->_height;
    		$this->_width = round($ratio * $pValue);
    	}

    	// Set height
    	$this->_height = $pValue;

    	return $this;
    }

    /**
     * Set width and height with proportional resize
     * @author Vincent@luo MSN:kele_100@hotmail.com
     * @param int $width
     * @param int $height
     * @example $objDrawing->setResizeProportional(true);
     * @example $objDrawing->setWidthAndHeight(160,120);
     * @return PHPPowerPoint_Shape_BaseDrawing
     */
	public function setWidthAndHeight($width = 0, $height = 0) {
		$xratio = $width / $this->_width;
		$yratio = $height / $this->_height;
		if ($this->_resizeProportional && !($width == 0 || $height == 0)) {
			if (($xratio * $this->_height) < $height) {
				$this->_height = ceil($xratio * $this->_height);
				$this->_width  = $width;
			} else {
				$this->_width	= ceil($yratio * $this->_width);
				$this->_height	= $height;
			}
		}
		return $this;
	}

    /**
     * Get ResizeProportional
     *
     * @return boolean
     */
    public function getResizeProportional() {
    	return $this->_resizeProportional;
    }

    /**
     * Set ResizeProportional
     *
     * @param boolean $pValue
     * @return PHPPowerPoint_Shape_BaseDrawing
     */
    public function setResizeProportional($pValue = true) {
    	$this->_resizeProportional = $pValue;
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
    		. $this->_description
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
