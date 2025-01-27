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
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

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
    public const LINE_SPACING_MODE_PERCENT = 'percent';
    public const LINE_SPACING_MODE_POINT = 'point';

    /**
     * Rich text elements.
     *
     * @var array<TextElementInterface>
     */
    private $richTextElements = [];

    /**
     * Alignment.
     *
     * @var Alignment
     */
    private $alignment;

    /**
     * Font.
     *
     * @var null|Font
     */
    private $font;

    /**
     * Bullet style.
     *
     * @var Bullet
     */
    private $bulletStyle;

    /**
     * @var int
     */
    private $lineSpacing = 100;

    /**
     * @var string
     */
    private $lineSpacingMode = self::LINE_SPACING_MODE_PERCENT;

    /**
     * @var float
     */
    private $spacingBefore = 0;

    /**
     * @var float
     */
    private $spacingAfter = 0;

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
        $this->alignment = new Alignment();
        $this->font = new Font();
        $this->bulletStyle = new Bullet();
    }

    /**
     * Magic Method : clone.
     */
    public function __clone()
    {
        // Clone each text
        foreach ($this->richTextElements as &$rtElement) {
            $rtElement = clone $rtElement;
        }
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
     * @param null|Font $pFont Font
     */
    public function setFont(?Font $pFont = null): self
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
     */
    public function setBulletStyle(?Bullet $style = null): self
    {
        $this->bulletStyle = $style;

        return $this;
    }

    /**
     * Create text (can not be formatted !).
     *
     * @param string $pText Text
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
     * @param null|TextElementInterface $pText Rich text element
     */
    public function addText(?TextElementInterface $pText = null): self
    {
        $this->richTextElements[] = $pText;

        return $this;
    }

    /**
     * Create break.
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
     * @return null|int Hash index
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

    public function getLineSpacing(): int
    {
        return $this->lineSpacing;
    }

    /**
     * Value in points.
     */
    public function setLineSpacing(int $lineSpacing): self
    {
        $this->lineSpacing = $lineSpacing;

        return $this;
    }

    public function getLineSpacingMode(): string
    {
        return $this->lineSpacingMode;
    }

    public function setLineSpacingMode(string $lineSpacingMode): self
    {
        if (in_array($lineSpacingMode, [
            self::LINE_SPACING_MODE_PERCENT,
            self::LINE_SPACING_MODE_POINT,
        ])) {
            $this->lineSpacingMode = $lineSpacingMode;
        }

        return $this;
    }

    /**
     * Value in points.
     */
    public function getSpacingBefore(): float
    {
        return $this->spacingBefore;
    }

    /**
     * Value in points.
     */
    public function setSpacingBefore(float $spacingBefore): self
    {
        $this->spacingBefore = $spacingBefore;

        return $this;
    }

    /**
     * Value in points.
     */
    public function getSpacingAfter(): float
    {
        return $this->spacingAfter;
    }

    /**
     * Value in points.
     */
    public function setSpacingAfter(float $spacingAfter): self
    {
        $this->spacingAfter = $spacingAfter;

        return $this;
    }
}
