<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPowerpoint\Shape;

use PhpOffice\PhpPowerpoint\AbstractShape;
use PhpOffice\PhpPowerpoint\ComparableInterface;
use PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph;
use PhpOffice\PhpPowerpoint\Shape\RichText\TextElementInterface;

/**
 * \PhpOffice\PhpPowerpoint\Shape\RichText
 */
class RichText extends AbstractShape implements ComparableInterface
{
    /** Wrapping */
    const WRAP_NONE = 'none';
    const WRAP_SQUARE = 'square';

    /** Autofit */
    const AUTOFIT_DEFAULT = 'spAutoFit';
    const AUTOFIT_SHAPE = 'spAutoFit';
    const AUTOFIT_NOAUTOFIT = 'noAutofit';
    const AUTOFIT_NORMAL = 'normAutoFit';

    /** Overflow */
    const OVERFLOW_CLIP = 'clip';
    const OVERFLOW_OVERFLOW = 'overflow';

    /**
     * Rich text paragraphs
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph[]
     */
    private $richTextParagraphs;

    /**
     * Active paragraph
     *
     * @var int
     */
    private $activeParagraph = 0;

    /**
     * Text wrapping
     *
     * @var string
     */
    private $wrap = self::WRAP_SQUARE;

    /**
     * Autofit
     *
     * @var string
     */
    private $autoFit = self::AUTOFIT_DEFAULT;

    /**
     * Horizontal overflow
     *
     * @var string
     */
    private $horizontalOverflow = self::OVERFLOW_OVERFLOW;

    /**
     * Vertical overflow
     *
     * @var string
     */
    private $verticalOverflow = self::OVERFLOW_OVERFLOW;

    /**
     * Text upright?
     *
     * @var boolean
     */
    private $upright = false;

    /**
     * Vertical text?
     *
     * @var boolean
     */
    private $vertical = false;

    /**
     * Number of columns (1 - 16)
     *
     * @var int
     */
    private $columns = 1;

    /**
     * Bottom inset (in pixels)
     *
     * @var float
     */
    private $bottomInset = 4.8;

    /**
     * Left inset (in pixels)
     *
     * @var float
     */
    private $leftInset = 9.6;

    /**
     * Right inset (in pixels)
     *
     * @var float
     */
    private $rightInset = 9.6;

    /**
     * Top inset (in pixels)
     *
     * @var float
     */
    private $topInset = 4.8;

    /**
     * Horizontal Auto Shrink
     * @var boolean
     */
    private $autoShrinkHorizontal;

    /**
     * Vertical Auto Shrink
     * @var boolean
     */
    private $autoShrinkVertical;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Shape\RichText instance
     */
    public function __construct()
    {
        // Initialise variables
        $this->richTextParagraphs = array(
            new Paragraph()
        );
        $this->activeParagraph    = 0;

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
        return $this->activeParagraph;
    }

    /**
     * Get active paragraph
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph
     */
    public function getActiveParagraph()
    {
        return $this->richTextParagraphs[$this->activeParagraph];
    }

    /**
     * Set active paragraph
     *
     * @param  int $index
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph
     */
    public function setActiveParagraph($index = 0)
    {
        if ($index >= count($this->richTextParagraphs)) {
            throw new \Exception("Invalid paragraph count.");
        }

        $this->activeParagraph = $index;

        return $this->getActiveParagraph();
    }

    /**
     * Get paragraph
     *
     * @param  int $index
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph
     */
    public function getParagraph($index = 0)
    {
        if ($index >= count($this->richTextParagraphs)) {
            throw new \Exception("Invalid paragraph count.");
        }

        return $this->richTextParagraphs[$index];
    }

    /**
     * Create paragraph
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph
     */
    public function createParagraph()
    {
        $alignment   = clone $this->getActiveParagraph()->getAlignment();
        $font        = clone $this->getActiveParagraph()->getFont();
        $bulletStyle = clone $this->getActiveParagraph()->getBulletStyle();

        $this->richTextParagraphs[] = new Paragraph();
        $this->activeParagraph      = count($this->richTextParagraphs) - 1;

        $this->getActiveParagraph()->setAlignment($alignment);
        $this->getActiveParagraph()->setFont($font);
        $this->getActiveParagraph()->setBulletStyle($bulletStyle);

        return $this->getActiveParagraph();
    }

    /**
     * Add text
     *
     * @param  \PhpOffice\PhpPowerpoint\Shape\RichText\TextElementInterface $pText Rich text element
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function addText(TextElementInterface $pText = null)
    {
        $this->richTextParagraphs[$this->activeParagraph]->addText($pText);

        return $this;
    }

    /**
     * Create text (can not be formatted !)
     *
     * @param  string                                   $pText Text
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\TextElement
     * @throws \Exception
     */
    public function createText($pText = '')
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createText($pText);
    }

    /**
     * Create break
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\BreakElement
     * @throws \Exception
     */
    public function createBreak()
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createBreak();
    }

    /**
     * Create text run (can be formatted)
     *
     * @param  string                           $pText Text
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\Run
     * @throws \Exception
     */
    public function createTextRun($pText = '')
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createTextRun($pText);
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

        // Loop trough all \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph
        foreach ($this->richTextParagraphs as $p) {
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
    public function __toString()
    {
        return $this->getPlainText();
    }

    /**
     * Get paragraphs
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph[]
     */
    public function getParagraphs()
    {
        return $this->richTextParagraphs;
    }

    /**
     * Set paragraphs
     *
     * @param  \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph[] $paragraphs Array of paragraphs
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setParagraphs($paragraphs = null)
    {
        if (is_array($paragraphs)) {
            $this->richTextParagraphs = $paragraphs;
            $this->activeParagraph    = count($this->richTextParagraphs) - 1;
        } else {
            throw new \Exception("Invalid \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph[] array passed.");
        }

        return $this;
    }

    /**
     * Get text wrapping
     *
     * @return string
     */
    public function getWrap()
    {
        return $this->wrap;
    }

    /**
     * Set text wrapping
     *
     * @param $value string
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setWrap($value = self::WRAP_SQUARE)
    {
        $this->wrap = $value;

        return $this;
    }

    /**
     * Get autofit
     *
     * @return string
     */
    public function getAutoFit()
    {
        return $this->autoFit;
    }

    /**
     * Set autofit
     *
     * @param $value string
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setAutoFit($value = self::AUTOFIT_DEFAULT)
    {
        $this->autoFit = $value;

        return $this;
    }

    /**
     * Get horizontal overflow
     *
     * @return string
     */
    public function getHorizontalOverflow()
    {
        return $this->horizontalOverflow;
    }

    /**
     * Set horizontal overflow
     *
     * @param $value string
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setHorizontalOverflow($value = self::OVERFLOW_OVERFLOW)
    {
        $this->horizontalOverflow = $value;

        return $this;
    }

    /**
     * Get vertical overflow
     *
     * @return string
     */
    public function getVerticalOverflow()
    {
        return $this->verticalOverflow;
    }

    /**
     * Set vertical overflow
     *
     * @param $value string
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setVerticalOverflow($value = self::OVERFLOW_OVERFLOW)
    {
        $this->verticalOverflow = $value;

        return $this;
    }

    /**
     * Get upright
     *
     * @return boolean
     */
    public function isUpright()
    {
        return $this->upright;
    }

    /**
     * Set vertical
     *
     * @param $value boolean
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setUpright($value = false)
    {
        $this->upright = $value;

        return $this;
    }

    /**
     * Get vertical
     *
     * @return boolean
     */
    public function isVertical()
    {
        return $this->vertical;
    }

    /**
     * Set vertical
     *
     * @param $value boolean
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setVertical($value = false)
    {
        $this->vertical = $value;

        return $this;
    }

    /**
     * Get columns
     *
     * @return int
     */
    public function getColumns()
    {
        return $this->columns;
    }

    /**
     * Set columns
     *
     * @param $value int
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setColumns($value = 1)
    {
        if ($value > 16 || $value < 1) {
            throw new \Exception('Number of columns should be 1-16');
        }

        $this->columns = $value;

        return $this;
    }

    /**
     * Get bottom inset
     *
     * @return float
     */
    public function getInsetBottom()
    {
        return $this->bottomInset;
    }

    /**
     * Set bottom inset
     *
     * @param $value float
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setInsetBottom($value = 4.8)
    {
        $this->bottomInset = $value;

        return $this;
    }

    /**
     * Get left inset
     *
     * @return float
     */
    public function getInsetLeft()
    {
        return $this->leftInset;
    }

    /**
     * Set left inset
     *
     * @param $value float
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setInsetLeft($value = 9.6)
    {
        $this->leftInset = $value;

        return $this;
    }

    /**
     * Get right inset
     *
     * @return float
     */
    public function getInsetRight()
    {
        return $this->rightInset;
    }

    /**
     * Set left inset
     *
     * @param $value float
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setInsetRight($value = 9.6)
    {
        $this->rightInset = $value;

        return $this;
    }

    /**
     * Get top inset
     *
     * @return float
     */
    public function getInsetTop()
    {
        return $this->topInset;
    }

    /**
     * Set top inset
     *
     * @param $value float
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
     */
    public function setInsetTop($value = 4.8)
    {
        $this->topInset = $value;

        return $this;
    }

    /**
     * Set horizontal auto shrink
     * @param bool $value
     */
    public function setAutoShrinkHorizontal($value = null)
    {
        if (is_bool($value)) {
            $this->autoShrinkHorizontal = $value;
        }
        return $this;
    }
    
    /**
     * Get horizontal auto shrink
     * @return bool
     */
    public function hasAutoShrinkHorizontal()
    {
        return $this->autoShrinkHorizontal;
    }
    
    /**
     * Set vertical auto shrink
     * @param bool $value
     */
    public function setAutoShrinkVertical($value = null)
    {
        if (is_bool($value)) {
            $this->autoShrinkVertical = $value;
        }
        return $this;
    }
    
    /**
     * Set vertical auto shrink
     * @return bool
     */
    public function hasAutoShrinkVertical()
    {
        return $this->autoShrinkVertical;
    }
    
    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        $hashElements = '';
        foreach ($this->richTextParagraphs as $element) {
            $hashElements .= $element->getHashCode();
        }

        return md5($hashElements . $this->wrap . $this->autoFit . $this->horizontalOverflow . $this->verticalOverflow . ($this->upright ? '1' : '0') . ($this->vertical ? '1' : '0') . $this->columns . $this->bottomInset . $this->leftInset . $this->rightInset . $this->topInset . parent::getHashCode() . __CLASS__);
    }
}
