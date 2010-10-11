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
 * PHPPowerPoint_Shape_RichText
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_RichText
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_RichText extends PHPPowerPoint_Shape implements PHPPowerPoint_IComparable
{
	/** Wrapping */
	const WRAP_NONE   = 'none';
	const WRAP_SQUARE = 'square';
	
	/** Autofit */
	const AUTOFIT_DEFAULT   = 'spAutoFit';
	const AUTOFIT_SHAPE     = 'spAutoFit';
	const AUTOFIT_NOAUTOFIT = 'noAutofit';
	const AUTOFIT_NORMAL    = 'normAutoFit';
	
	/** Overflow */
	const OVERFLOW_CLIP     = 'clip';
	const OVERFLOW_OVERFLOW = 'overflow';
	
	/**
	 * Rich text paragraphs
	 *
	 * @var PHPPowerPoint_Shape_RichText_Paragraph[]
	 */
	private $_richTextParagraphs;

	/**
	 * Active paragraph
	 *
	 * @var int
	 */
	private $_activeParagraph = 0;
	
	/**
	 * Text wrapping
	 * 
	 * @var string
	 */
	private $_wrap = PHPPowerPoint_Shape_RichText::WRAP_SQUARE;
	
	/**
	 * Autofit
	 * 
	 * @var string
	 */
	private $_autoFit = PHPPowerPoint_Shape_RichText::AUTOFIT_DEFAULT;
	
	/**
	 * Horizontal overflow
	 * 
	 * @var string
	 */
	private $_horizontalOverflow = PHPPowerPoint_Shape_RichText::OVERFLOW_OVERFLOW;
	
	/**
	 * Vertical overflow
	 * 
	 * @var string
	 */
	private $_verticalOverflow = PHPPowerPoint_Shape_RichText::OVERFLOW_OVERFLOW;

	/**
	 * Text upright?
	 * 
	 * @var boolean
	 */
	private $_upright = false;
	
	/**
	 * Vertical text?
	 * 
	 * @var boolean
	 */
	private $_vertical = false;
	
	/**
	 * Number of columns (1 - 16)
	 * 
	 * @var int
	 */
	private $_columns = 1;
	
	/**
	 * Bottom inset (in pixels)
	 * 
	 * @var float
	 */
	private $_bottomInset = 4.8;
	
	/**
	 * Left inset (in pixels)
	 * 
	 * @var float
	 */
	private $_leftInset = 9.6;
	
	/**
	 * Right inset (in pixels)
	 * 
	 * @var float
	 */
	private $_rightInset = 9.6;
	
	/**
	 * Top inset (in pixels)
	 * 
	 * @var float
	 */
	private $_topInset = 4.8;

    /**
     * Create a new PHPPowerPoint_Shape_RichText instance
     */
    public function __construct()
    {
    	// Initialise variables
    	$this->_richTextParagraphs = array(
    		new PHPPowerPoint_Shape_RichText_Paragraph()
    	);
    	$this->_activeParagraph = 0;

    	// Initialize parent
    	parent::__construct();
    }

    /**
     * Get active paragraph index
     *
     * @return int
     */
	public function getActiveParagraphIndex()
    {
    	return $this->_activeParagraph;
    }

    /**
     * Get active paragraph
     *
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
	public function getActiveParagraph()
    {
    	return $this->_richTextParagraphs[$this->_activeParagraph];
    }

    /**
     * Set active paragraph
     *
     * @param int $index
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
	public function setActiveParagraph($index = 0)
    {
    	if ($index >= count($this->_richTextParagraphs))
    		throw new Exception("Invalid paragraph count.");

    	$this->_activeParagraph = $index;
    	return $this->getActiveParagraph();
    }

    /**
     * Get paragraph
     *
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
	public function getParagraph($index = 0)
    {
    	if ($index >= count($this->_richTextParagraphs))
    		throw new Exception("Invalid paragraph count.");

    	return $this->_richTextParagraphs[$index];
    }

    /**
     * Create paragraph
     *
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
	public function createParagraph()
    {
	    $alignment = clone $this->getActiveParagraph()->getAlignment();
	    $font = clone $this->getActiveParagraph()->getFont();
	    $bulletStyle = clone $this->getActiveParagraph()->getBulletStyle();

    	$this->_richTextParagraphs[] = new PHPPowerPoint_Shape_RichText_Paragraph();
    	$this->_activeParagraph = count($this->_richTextParagraphs) - 1;

    	$this->getActiveParagraph()->setAlignment($alignment);
    	$this->getActiveParagraph()->setFont($font);
    	$this->getActiveParagraph()->setBulletStyle($bulletStyle);

    	return $this->getActiveParagraph();
    }

    /**
     * Add text
     *
     * @param 	PHPPowerPoint_Shape_RichText_ITextElement		$pText		Rich text element
     * @throws 	Exception
     * @return PHPPowerPoint_Shape_RichText
     */
    public function addText(PHPPowerPoint_Shape_RichText_ITextElement $pText = null)
    {
    	$this->_richTextParagraphs[$this->_activeParagraph]->addText($pText);
    	return $this;
    }

    /**
     * Create text (can not be formatted !)
     *
     * @param 	string	$pText	Text
     * @return	PHPPowerPoint_Shape_RichText_TextElement
     * @throws 	Exception
     */
    public function createText($pText = '')
    {
    	return $this->_richTextParagraphs[$this->_activeParagraph]->createText($pText);
    }

    /**
     * Create break
     *
     * @return	PHPPowerPoint_Shape_RichText_Break
     * @throws 	Exception
     */
    public function createBreak()
    {
    	return $this->_richTextParagraphs[$this->_activeParagraph]->createBreak();
    }

    /**
     * Create text run (can be formatted)
     *
     * @param 	string	$pText	Text
     * @return	PHPPowerPoint_Shape_RichText_Run
     * @throws 	Exception
     */
    public function createTextRun($pText = '')
    {
    	return $this->_richTextParagraphs[$this->_activeParagraph]->createTextRun($pText);
    }

    /**
     * Get plain text
     *
     * @return string
     */
    public function getPlainText()
    {
    	// Return value
    	$returnValue = '';

    	// Loop trough all PHPPowerPoint_Shape_RichText_Paragraph
    	foreach ($this->_richTextParagraphs as $p) {
    		$returnValue .= $p->getPlainText();
    	}

    	// Return
    	return $returnValue;
    }

    /**
     * Convert to string
     *
     * @return string
     */
    public function __toString() {
    	return $this->getPlainText();
    }

    /**
     * Get paragraphs
     *
     * @return PHPPowerPoint_Shape_RichText_Paragraph[]
     */
    public function getParagraphs()
    {
    	return $this->_richTextParagraphs;
    }

    /**
     * Set paragraphs
     *
     * @param 	PHPPowerPoint_Shape_RichText_Paragraphs[]	$paragraphs		Array of paragraphs
     * @throws 	Exception
     * @return PHPPowerPoint_Shape_RichText
     */
    public function setParagraphs($paragraphs = null)
    {
    	if (is_array($paragraphs)) {
    		$this->_richTextParagraphs = $paragraphs;
    		$this->_activeParagraph = count($this->_richTextParagraphs) - 1;
    	} else {
    		throw new Exception("Invalid PHPPowerPoint_Shape_RichText_Paragraph[] array passed.");
    	}
    	return $this;
    }
    
    /**
     * Get text wrapping
     * 
     * @return string
     */
    public function getWrap() {
    	return $this->_wrap;
    }
    
    /**
     * Set text wrapping
     * 
     * @param $value string
     * @return PHPPowerPoint_Shape_RichText
     */
    public function setWrap($value = PHPPowerPoint_Shape_RichText::WRAP_SQUARE) {
		$this->_wrap = $value;
		return $this;
    }
    
    /**
     * Get autofit
     * 
     * @return string
     */
    public function getAutoFit() {
    	return $this->_autoFit;
    }
    
    /**
     * Set autofit
     * 
     * @param $value string
     * @return PHPPowerPoint_Shape_RichText
     */
    public function setAutoFit($value = PHPPowerPoint_Shape_RichText::AUTOFIT_DEFAULT) {
		$this->_autoFit = $value;
		return $this;
    }
    
    /**
     * Get horizontal overflow
     * 
     * @return string
     */
	public function getHorizontalOverflow() {
    	return $this->_horizontalOverflow;
    }
    
    /**
     * Set horizontal overflow
     * 
     * @param $value string
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setHorizontalOverflow($value = PHPPowerPoint_Shape_RichText::OVERFLOW_OVERFLOW) {
    	$this->_horizontalOverflow = $value;
    	return $this;
    }
    
    /**
     * Get vertical overflow
     * 
     * @return string
     */
	public function getVerticalOverflow() {
    	return $this->_verticalOverflow;
    }
    
    /**
     * Set vertical overflow
     * 
     * @param $value string
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setVerticalOverflow($value = PHPPowerPoint_Shape_RichText::OVERFLOW_OVERFLOW) {
    	$this->_verticalOverflow = $value;
    	return $this;
    }
    
    /**
     * Get upright
     * 
     * @return boolean
     */
	public function getUpright() {
    	return $this->_upright;
    }
    
    /**
     * Set vertical
     * 
     * @param $value boolean
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setUpright($value = false) {
    	$this->_upright = $value;
    	return $this;
    }
    
    /**
     * Get vertical
     * 
     * @return boolean
     */
	public function getVertical() {
    	return $this->_vertical;
    }
    
    /**
     * Set vertical
     * 
     * @param $value boolean
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setVertical($value = false) {
    	$this->_vertical = $value;
    	return $this;
    }
    
    /**
     * Get columns
     * 
     * @return int
     */
	public function getColumns() {
    	return $this->_columns;
    }
    
    /**
     * Set columns
     * 
     * @param $value int
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setColumns($value = 1) {
		if ($value > 16 || $value < 1) {
			throw new Exception('Number of columns should be 1-16');
		}
		
    	$this->_columns = $value;
    	return $this;
    }
    
    /**
     * Get bottom inset
     * 
     * @return float
     */
	public function getInsetBottom() {
    	return $this->_bottomInset;
    }
    
    /**
     * Set bottom inset
     * 
     * @param $value float
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setInsetBottom($value = 4.8) {
    	$this->_bottomInset = $value;
    	return $this;
    }
    
    /**
     * Get left inset
     * 
     * @return float
     */
	public function getInsetLeft() {
    	return $this->_leftInset;
    }
    
    /**
     * Set left inset
     * 
     * @param $value float
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setInsetLeft($value = 9.6) {
    	$this->_leftInset = $value;
    	return $this;
    }
    
    /**
     * Get right inset
     * 
     * @return float
     */
	public function getInsetRight() {
    	return $this->_rightInset;
    }
    
    /**
     * Set left inset
     * 
     * @param $value float
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setInsetRight($value = 9.6) {
    	$this->_rightInset = $value;
    	return $this;
    }
    
    /**
     * Get top inset
     * 
     * @return float
     */
	public function getInsetTop() {
    	return $this->_topInset;
    }
    
    /**
     * Set top inset
     * 
     * @param $value float
     * @return PHPPowerPoint_Shape_RichText
     */
	public function setInsetTop($value = 4.8) {
    	$this->_topInset = $value;
    	return $this;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
		$hashElements = '';
		foreach ($this->_richTextParagraphs as $element) {
			$hashElements .= $element->getHashCode();
		}

    	return md5(
    		  $hashElements
    		. $this->_wrap
    		. $this->_autoFit
    		. $this->_horizontalOverflow
    		. $this->_verticalOverflow
    		. ($this->_upright ? '1' : '0')
    		. ($this->_vertical ? '1' : '0')
    		. $this->_columns
    		. $this->_bottomInset
    		. $this->_leftInset
    		. $this->_rightInset
    		. $this->_topInset
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
			if ($key == '_parent') continue;

			if (is_object($value)) {
				$this->$key = clone $value;
			} else {
				$this->$key = $value;
			}
		}
	}
}
