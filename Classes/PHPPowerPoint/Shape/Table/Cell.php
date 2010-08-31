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
 * PHPPowerPoint_Shape_Table_Cell
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Shape_Table_Cell implements PHPPowerPoint_IComparable
{
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
	 * Fill
	 *
	 * @var PHPPowerPoint_Style_Fill
	 */
	private $_fill;

	/**
	 * Borders
	 *
	 * @var PHPPowerPoint_Style_Borders
	 */
	private $_borders;

	/**
	 * Width (in pixels)
	 *
	 * @var int
	 */
	private $_width = 0;

	/**
	 * Colspan
	 *
	 * @var int
	 */
	private $_colSpan = 0;

	/**
	 * Rowspan
	 *
	 * @var int
	 */
	private $_rowSpan = 0;

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

    	// Set fill
    	$this->_fill = new PHPPowerPoint_Style_Fill();

    	// Set borders
    	$this->_borders = new PHPPowerPoint_Style_Borders();
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
     * Get fill
     *
     * @return PHPPowerPoint_Style_Fill
     */
    public function getFill()
    {
    	return $this->_fill;
    }

    /**
     * Set fill
     *
     * @param PHPPowerPoint_Style_Fill $fill
     * @return PHPPowerPoint_Shape_RichText
     */
    public function setFill(PHPPowerPoint_Style_Fill $fill)
    {
    	$this->_fill = $fill;
    	return $this;
    }

    /**
     * Get borders
     *
     * @return PHPPowerPoint_Style_Borders
     */
    public function getBorders()
    {
    	return $this->_borders;
    }

    /**
     * Set borders
     *
     * @param PHPPowerPoint_Style_Borders $borders
     * @return PHPPowerPoint_Shape_RichText
     */
    public function setBorders(PHPPowerPoint_Style_Borders $borders)
    {
    	$this->_borders = $borders;
    	return $this;
    }

    /**
     * Get width
     *
     * @return int
     */
    public function getWidth()
    {
    	return $this->_width;
    }

    /**
     * Set width
     *
     * @param int $value
     * @return PHPPowerPoint_Shape_RichText
     */
    public function setWidth($value = 0)
    {
    	$this->_width = $value;
    	return $this;
    }

    /**
     * Get colSpan
     *
     * @return int
     */
    public function getColSpan()
    {
    	return $this->_colSpan;
    }

    /**
     * Set colSpan
     *
     * @param int $value
     * @return PHPPowerPoint_Shape_RichText
     */
    public function setColSpan($value = 0)
    {
    	$this->_colSpan = $value;
    	return $this;
    }

    /**
     * Get rowSpan
     *
     * @return int
     */
    public function getRowSpan()
    {
    	return $this->_rowSpan;
    }

    /**
     * Set rowSpan
     *
     * @param int $value
     * @return PHPPowerPoint_Shape_RichText
     */
    public function setRowSpan($value = 0)
    {
    	$this->_rowSpan = $value;
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
    		. $this->_fill->getHashCode()
    		. $this->_borders->getHashCode()
    		. $this->_width
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
