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
 * PHPPowerPoint_Shape
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
abstract class PHPPowerPoint_Shape implements PHPPowerPoint_IComparable
{
	/**
	 * Slide
	 *
	 * @var PHPPowerPoint_Slide
	 */
	protected $_slide;

	/**
	 * Offset X
	 *
	 * @var int
	 */
	protected $_offsetX;

	/**
	 * Offset Y
	 *
	 * @var int
	 */
	protected $_offsetY;

	/**
	 * Width
	 *
	 * @var int
	 */
	protected $_width;

	/**
	 * Height
	 *
	 * @var int
	 */
	protected $_height;

	/**
	 * Fill
	 *
	 * @var PHPPowerPoint_Style_Fill
	 */
	private $_fill;

	/**
	 * Border
	 *
	 * @var PHPPowerPoint_Style_Border
	 */
	private $_border;

	/**
	 * Rotation
	 *
	 * @var int
	 */
	protected $_rotation;

	/**
	 * Shadow
	 *
	 * @var PHPPowerPoint_Shape_Shadow
	 */
	protected $_shadow;
	
	/**
	 * Hyperlink
	 * 
	 * @var PHPPowerPoint_Shape_Hyperlink
	 */
	protected $_hyperlink;

    /**
     * Create a new PHPPowerPoint_Shape
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_slide				= null;
    	$this->_offsetX				= 0;
    	$this->_offsetY				= 0;
    	$this->_width				= 0;
    	$this->_height				= 0;
    	$this->_rotation			= 0;
    	$this->_fill				= new PHPPowerPoint_Style_Fill();
    	$this->_border				= new PHPPowerPoint_Style_Border();
    	$this->_shadow				= new PHPPowerPoint_Shape_Shadow();
    	
    	$this->_border->setLineStyle(PHPPowerPoint_Style_Border::LINE_NONE);
    }

    /**
     * Get Slide
     *
     * @return PHPPowerPoint_Slide
     */
    public function getSlide() {
    	return $this->_slide;
    }

    /**
     * Set Slide
     *
     * @param 	PHPPowerPoint_Slide 	$pValue
     * @param 	bool					$pOverrideOld	If a Slide has already been assigned, overwrite it and remove image from old Slide?
     * @throws 	Exception
     * @return PHPPowerPoint_Shape
     */
    public function setSlide(PHPPowerPoint_slide $pValue = null, $pOverrideOld = false) {
    	if (is_null($this->_slide)) {
    		// Add drawing to PHPPowerPoint_Slide
	    	$this->_slide = $pValue;
	    	$this->_slide->getShapeCollection()->append($this);
    	} else {
    		if ($pOverrideOld) {
    			// Remove drawing from old PHPPowerPoint_Slide
    			$iterator = $this->_slide->getShapeCollection()->getIterator();

    			while ($iterator->valid()) {
    				if ($iterator->current()->getHashCode() == $this->getHashCode()) {
    					$this->_slide->getShapeCollection()->offsetUnset( $iterator->key() );
    					$this->_slide = null;
    					break;
    				}
    			}

    			// Set new PHPPowerPoint_Slide
    			$this->setSlide($pValue);
    		} else {
    			throw new Exception("A PHPPowerPoint_Slide has already been assigned. Shapes can only exist on one PHPPowerPoint_Slide.");
    		}
    	}
    	return $this;
    }

    /**
     * Get OffsetX
     *
     * @return int
     */
    public function getOffsetX() {
    	return $this->_offsetX;
    }

    /**
     * Set OffsetX
     *
     * @param int $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setOffsetX($pValue = 0) {
    	$this->_offsetX = $pValue;
    	return $this;
    }

    /**
     * Get OffsetY
     *
     * @return int
     */
    public function getOffsetY() {
    	return $this->_offsetY;
    }

    /**
     * Set OffsetY
     *
     * @param int $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setOffsetY($pValue = 0) {
    	$this->_offsetY = $pValue;
    	return $this;
    }

    /**
     * Get Width
     *
     * @return int
     */
    public function getWidth() {
    	return $this->_width;
    }

    /**
     * Set Width
     *
     * @param int $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setWidth($pValue = 0) {
    	$this->_width = $pValue;
    	return $this;
    }

    /**
     * Get Height
     *
     * @return int
     */
    public function getHeight() {
    	return $this->_height;
    }

    /**
     * Set Height
     *
     * @param int $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setHeight($pValue = 0) {
    	$this->_height = $pValue;
    	return $this;
    }

    /**
     * Set width and height with proportional resize
     *
     * @param int $width
     * @param int $height
     * @example $objDrawing->setWidthAndHeight(160,120);
     * @return PHPPowerPoint_Shape
     */
	public function setWidthAndHeight($width = 0, $height = 0) {
		$this->_width  = $width;
		$this->_height	= $height;
		return $this;
	}

    /**
     * Get Rotation
     *
     * @return int
     */
    public function getRotation() {
    	return $this->_rotation;
    }

    /**
     * Set Rotation
     *
     * @param int $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setRotation($pValue = 0) {
    	$this->_rotation = $pValue;
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
     * Get Border
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getBorder() {
		return $this->_border;
    }

    /**
     * Get Shadow
     *
     * @return PHPPowerPoint_Shape_Shadow
     */
    public function getShadow() {
    	return $this->_shadow;
    }

    /**
     * Set Shadow
     *
     * @param 	PHPPowerPoint_Shape_Shadow $pValue
     * @throws 	Exception
     * @return PHPPowerPoint_Shape
     */
    public function setShadow(PHPPowerPoint_Shape_Shadow $pValue = null) {
   		$this->_shadow = $pValue;
   		return $this;
    }
    
	/**
	 * Has Hyperlink?
	 *
	 * @return boolean
	 */
	public function hasHyperlink()
	{
		return !is_null($this->_hyperlink);
	}
    
	/**
	 * Get Hyperlink
	 *
	 * @return PHPPowerPoint_Shape_Hyperlink
	 */
	public function getHyperlink()
	{
		if (is_null($this->_hyperlink)) {
			$this->_hyperlink = new PHPPowerPoint_Shape_Hyperlink();
		}
		return $this->_hyperlink;
	}

	/**
	 * Set Hyperlink
	 *
	 * @param	PHPPowerPoint_Shape_Hyperlink	$pHyperlink
	 * @throws	Exception
	 * @return PHPPowerPoint_Shape
	 */
	public function setHyperlink(PHPExcel_Cell_Hyperlink $pHyperlink = null)
	{
		$this->_hyperlink = $pHyperlink;
		return $this;
	}

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_slide->getHashCode()
    		. $this->_offsetX
    		. $this->_offsetY
    		. $this->_width
    		. $this->_height
    		. $this->_rotation
    		. $this->getFill()->getHashCode()
    		. $this->_shadow->getHashCode()
    		. (is_null($this->_hyperlink) ? '' : $this->_hyperlink->getHashCode())
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
