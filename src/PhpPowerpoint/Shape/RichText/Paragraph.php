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

namespace PhpOffice\PhpPowerpoint\Shape\RichText;

use PhpOffice\PhpPowerpoint\IComparable;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Font;
use PhpOffice\PhpPowerpoint\Style\Bullet;
use PhpOffice\PhpPowerpoint\Shape\RichText\TextElement;
use PhpOffice\PhpPowerpoint\Shape\RichText\BreakElement;

/**
 * PHPPowerPoint_Shape_RichText_Paragraph
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_RichText
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Paragraph implements IComparable
{
    /**
     * Rich text elements
     *
     * @var PHPPowerPoint_Shape_RichText_ITextElement[]
     */
    private $richTextElements;

    /**
     * Alignment
     *
     * @var PHPPowerPoint_Style_Alignment
     */
    private $alignment;

    /**
     * Font
     *
     * @var PHPPowerPoint_Style_Font
     */
    private $font;

    /**
     * Bullet style
     *
     * @var PHPPowerPoint_Style_Bullet
     */
    private $bulletStyle;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new PHPPowerPoint_Shape_RichText_Paragraph instance
     */
    public function __construct()
    {
        // Initialise variables
        $this->richTextElements = array();
        $this->alignment        = new Alignment();
        $this->font             = new Font();
        $this->bulletStyle      = new Bullet();
    }

    /**
     * Get alignment
     *
     * @return PHPPowerPoint_Style_Alignment
     */
    public function getAlignment()
    {
        return $this->alignment;
    }

    /**
     * Set alignment
     *
     * @param  PHPPowerPoint_Style_Alignment          $alignment
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
    public function setAlignment(Alignment $alignment)
    {
        $this->alignment = $alignment;

        return $this;
    }

    /**
     * Get font
     *
     * @return PHPPowerPoint_Style_Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param  PHPPowerPoint_Style_Font               $pFont Font
     * @throws \Exception
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
    public function setFont(Font $pFont = null)
    {
        $this->font = $pFont;

        return $this;
    }

    /**
     * Get bullet style
     *
     * @return PHPPowerPoint_Style_Bullet
     */
    public function getBulletStyle()
    {
        return $this->bulletStyle;
    }

    /**
     * Set bullet style
     *
     * @param  PHPPowerPoint_Style_Bullet             $style
     * @throws \Exception
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
    public function setBulletStyle(Bullet $style = null)
    {
        $this->bulletStyle = $style;

        return $this;
    }

    /**
     * Add text
     *
     * @param  PHPPowerPoint_Shape_RichText_ITextElement $pText Rich text element
     * @throws \Exception
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
    public function addText(ITextElement $pText = null)
    {
        $this->richTextElements[] = $pText;

        return $this;
    }

    /**
     * Create text (can not be formatted !)
     *
     * @param  string                                   $pText Text
     * @return PHPPowerPoint_Shape_RichText_TextElement
     * @throws \Exception
     */
    public function createText($pText = '')
    {
        $objText = new TextElement($pText);
        $this->addText($objText);

        return $objText;
    }

    /**
     * Create break
     *
     * @return PHPPowerPoint_Shape_RichText_Break
     * @throws \Exception
     */
    public function createBreak()
    {
        $objText = new BreakElement();
        $this->addText($objText);

        return $objText;
    }

    /**
     * Create text run (can be formatted)
     *
     * @param  string                           $pText Text
     * @return PHPPowerPoint_Shape_RichText_Run
     * @throws \Exception
     */
    public function createTextRun($pText = '')
    {
        $objText = new Run($pText);
        $objText->setFont(clone $this->font);
        $this->addText($objText);

        return $objText;
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

        // Loop trough all PHPPowerPoint_Shape_RichText_ITextElement
        foreach ($this->richTextElements as $text) {
            $returnValue .= $text->getText();
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
     * Get Rich Text elements
     *
     * @return PHPPowerPoint_Shape_RichText_ITextElement[]
     */
    public function getRichTextElements()
    {
        return $this->richTextElements;
    }

    /**
     * Set Rich Text elements
     *
     * @param  PHPPowerPoint_Shape_RichText_ITextElement[] $pElements Array of elements
     * @throws \Exception
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
    public function setRichTextElements($pElements = null)
    {
        if (is_array($pElements)) {
            $this->richTextElements = $pElements;
        } else {
            throw new \Exception("Invalid PHPPowerPoint_Shape_RichText_ITextElement[] array passed.");
        }

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        $hashElements = '';
        foreach ($this->richTextElements as $element) {
            $hashElements .= $element->getHashCode();
        }

        return md5($hashElements . $this->font->getHashCode() . __CLASS__);
    }

    /**
     * Get hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return string Hash index
     */
    public function getHashIndex()
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param string $value Hash index
     */
    public function setHashIndex($value)
    {
        $this->hashIndex = $value;
    }
}
