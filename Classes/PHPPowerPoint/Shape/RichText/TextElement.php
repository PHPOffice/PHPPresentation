<?php
/**
 * PHPPowerPoint
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
 * @package    PHPPowerPoint_RichText
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Shape_RichText_TextElement
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_RichText_TextElement implements PHPPowerPoint_Shape_RichText_ITextElement
{
	/**
	 * Text
	 *
	 * @var string
	 */
	private $_text;
	
	/**
	 * Hyperlink
	 * 
	 * @var PHPPowerPoint_Shape_Hyperlink
	 */
	protected $_hyperlink;

    /**
     * Create a new PHPPowerPoint_Shape_RichText_TextElement instance
     *
     * @param 	string		$pText		Text
     */
    public function __construct($pText = '')
    {
    	// Initialise variables
    	$this->_text = $pText;
    }

	/**
	 * Get text
	 *
	 * @return string	Text
	 */
	public function getText() {
		return $this->_text;
	}

	/**
	 * Set text
	 *
	 * @param 	$pText string	Text
	 * @return PHPPowerPoint_Shape_RichText_ITextElement
	 */
	public function setText($pText = '') {
		$this->_text = $pText;
		return $this;
	}

	/**
	 * Get font
	 *
	 * @return PHPPowerPoint_Style_Font
	 */
	public function getFont() {
		return null;
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
    		  $this->_text
    		. (is_null($this->_hyperlink) ? '' : $this->_hyperlink->getHashCode())
    		. __CLASS__
    	);
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
