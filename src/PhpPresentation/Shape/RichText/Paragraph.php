<?php
/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Shape\RichText;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Font;

/**
 * \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
 */
class Paragraph implements ComparableInterface
{
    /**
     * Rich text elements
     *
     * @var \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface[]
     */
    private $richTextElements;

    /**
     * Alignment
     *
     * @var \PhpOffice\PhpPresentation\Style\Alignment
     */
    private $alignment;

    /**
     * Font
     *
     * @var \PhpOffice\PhpPresentation\Style\Font
     */
    private $font;

    /**
     * Bullet style
     *
     * @var \PhpOffice\PhpPresentation\Style\Bullet
     */
    private $bulletStyle;

    /**
     * @var integer
     */
    private $lineSpacing = 100;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\RichText\Paragraph instance
     */
    public function __construct()
    {
        // Initialise variables
        $this->richTextElements = array();
        $this->alignment = new Alignment();
        $this->font = new Font();
        $this->bulletStyle = new Bullet();
    }

    /**
     * Get alignment
     *
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function getAlignment()
    {
        return $this->alignment;
    }

    /**
     * Set alignment
     *
     * @param  \PhpOffice\PhpPresentation\Style\Alignment $alignment
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
     */
    public function setAlignment(Alignment $alignment)
    {
        $this->alignment = $alignment;

        return $this;
    }

    /**
     * Get font
     *
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param  \PhpOffice\PhpPresentation\Style\Font $pFont Font
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
     */
    public function setFont(Font $pFont = null)
    {
        $this->font = $pFont;

        return $this;
    }

    /**
     * Get bullet style
     *
     * @return \PhpOffice\PhpPresentation\Style\Bullet
     */
    public function getBulletStyle()
    {
        return $this->bulletStyle;
    }

    /**
     * Set bullet style
     *
     * @param  \PhpOffice\PhpPresentation\Style\Bullet $style
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
     */
    public function setBulletStyle(Bullet $style = null)
    {
        $this->bulletStyle = $style;

        return $this;
    }

    /**
     * Create text (can not be formatted !)
     *
     * @param  string $pText Text
     * @return \PhpOffice\PhpPresentation\Shape\RichText\TextElement
     * @throws \Exception
     */
    public function createText($pText = '')
    {
        $objText = new TextElement($pText);
        $this->addText($objText);

        return $objText;
    }

    /**
     * Add text
     *
     * @param  \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface $pText Rich text element
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
     */
    public function addText(TextElementInterface $pText = null)
    {
        $this->richTextElements[] = $pText;

        return $this;
    }

    /**
     * Create break
     *
     * @return \PhpOffice\PhpPresentation\Shape\RichText\BreakElement
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
     * @param  string $pText Text
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Run
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
     * Convert to string
     *
     * @return string
     */
    public function __toString()
    {
        return $this->getPlainText();
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

        // Loop trough all \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface
        foreach ($this->richTextElements as $text) {
            if ($text instanceof TextElementInterface) {
                $returnValue .= $text->getText();
            }
        }

        // Return
        return $returnValue;
    }

    /**
     * Get Rich Text elements
     *
     * @return \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface[]
     */
    public function getRichTextElements()
    {
        return $this->richTextElements;
    }

    /**
     * Set Rich Text elements
     *
     * @param  \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface[] $pElements Array of elements
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
     */
    public function setRichTextElements($pElements = null)
    {
        if (is_array($pElements)) {
            $this->richTextElements = $pElements;
        } else {
            throw new \Exception("Invalid \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface[] array passed.");
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

    /**
     * @return int
     */
    public function getLineSpacing()
    {
        return $this->lineSpacing;
    }

    /**
     * @param int $lineSpacing
     * @return Paragraph
     */
    public function setLineSpacing($lineSpacing)
    {
        $this->lineSpacing = $lineSpacing;
        return $this;
    }
}
