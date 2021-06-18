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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Shape\RichText;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Font;

/**
 * \PhpOffice\PhpPresentation\Shape\RichText\Paragraph.
 */
class Paragraph implements ComparableInterface
{
    /**
     * Rich text elements.
     *
     * @var array<TextElementInterface>
     */
    private $richTextElements;

    /**
     * Alignment.
     *
     * @var Alignment
     */
    private $alignment;

    /**
     * Font.
     *
     * @var Font|null
     */
    private $font;

    /**
     * Bullet style.
     *
     * @var \PhpOffice\PhpPresentation\Style\Bullet
     */
    private $bulletStyle;

    /**
     * @var int
     */
    private $lineSpacing = 100;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\RichText\Paragraph instance.
     */
    public function __construct()
    {
        $this->richTextElements = [];
        $this->alignment = new Alignment();
        $this->font = new Font();
        $this->bulletStyle = new Bullet();
    }

    /**
     * Get alignment.
     */
    public function getAlignment(): Alignment
    {
        return $this->alignment;
    }

    /**
     * Set alignment.
     */
    public function setAlignment(Alignment $alignment): self
    {
        $this->alignment = $alignment;

        return $this;
    }

    /**
     * Get font.
     */
    public function getFont(): ?Font
    {
        return $this->font;
    }

    /**
     * Set font.
     *
     * @param Font|null $pFont Font
     *
     * @throws \Exception
     */
    public function setFont(Font $pFont = null): self
    {
        $this->font = $pFont;

        return $this;
    }

    /**
     * Get bullet style.
     */
    public function getBulletStyle(): ?Bullet
    {
        return $this->bulletStyle;
    }

    /**
     * Set bullet style.
     *
     * @throws \Exception
     */
    public function setBulletStyle(Bullet $style = null): self
    {
        $this->bulletStyle = $style;

        return $this;
    }

    /**
     * Create text (can not be formatted !).
     *
     * @param string $pText Text
     *
     * @throws \Exception
     */
    public function createText(string $pText = ''): TextElement
    {
        $objText = new TextElement($pText);
        $this->addText($objText);

        return $objText;
    }

    /**
     * Add text.
     *
     * @param TextElementInterface|null $pText Rich text element
     *
     * @throws \Exception
     */
    public function addText(TextElementInterface $pText = null): self
    {
        $this->richTextElements[] = $pText;

        return $this;
    }

    /**
     * Create break.
     *
     * @throws \Exception
     */
    public function createBreak(): BreakElement
    {
        $objText = new BreakElement();
        $this->addText($objText);

        return $objText;
    }

    /**
     * Create text run (can be formatted).
     *
     * @param string $pText Text
     *
     * @throws \Exception
     */
    public function createTextRun(string $pText = ''): Run
    {
        $objText = new Run($pText);
        $objText->setFont(clone $this->font);
        $this->addText($objText);

        return $objText;
    }

    /**
     * Convert to string.
     *
     * @return string
     */
    public function __toString()
    {
        return $this->getPlainText();
    }

    /**
     * Get plain text.
     */
    public function getPlainText(): string
    {
        // Return value
        $returnValue = '';

        // Loop trough all TextElementInterface
        foreach ($this->richTextElements as $text) {
            if ($text instanceof TextElementInterface) {
                $returnValue .= $text->getText();
            }
        }

        // Return
        return $returnValue;
    }

    /**
     * Get Rich Text elements.
     *
     * @return array<TextElementInterface>
     */
    public function getRichTextElements(): array
    {
        return $this->richTextElements;
    }

    /**
     * Set Rich Text elements.
     *
     * @param array<TextElementInterface> $pElements Array of elements
     */
    public function setRichTextElements(array $pElements = []): self
    {
        $this->richTextElements = $pElements;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        $hashElements = '';
        foreach ($this->richTextElements as $element) {
            $hashElements .= $element->getHashCode();
        }

        return md5($hashElements . $this->font->getHashCode() . __CLASS__);
    }

    /**
     * Get hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return int|null Hash index
     */
    public function getHashIndex(): ?int
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param int $value Hash index
     *
     * @return $this
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;

        return $this;
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
     *
     * @return Paragraph
     */
    public function setLineSpacing($lineSpacing)
    {
        $this->lineSpacing = $lineSpacing;

        return $this;
    }
}
