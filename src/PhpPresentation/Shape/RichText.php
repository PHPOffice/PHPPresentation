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

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Shape;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface;

/**
 * \PhpOffice\PhpPresentation\Shape\RichText.
 */
class RichText extends AbstractShape implements ComparableInterface
{
    /** Wrapping */
    public const WRAP_NONE = 'none';
    public const WRAP_SQUARE = 'square';

    /** Autofit */
    public const AUTOFIT_DEFAULT = 'spAutoFit';
    public const AUTOFIT_SHAPE = 'spAutoFit';
    public const AUTOFIT_NOAUTOFIT = 'noAutofit';
    public const AUTOFIT_NORMAL = 'normAutofit';

    /** Overflow */
    public const OVERFLOW_CLIP = 'clip';
    public const OVERFLOW_OVERFLOW = 'overflow';

    /**
     * Rich text paragraphs.
     *
     * @var array<Paragraph>
     */
    private $richTextParagraphs;

    /**
     * Active paragraph.
     *
     * @var int
     */
    private $activeParagraph = 0;

    /**
     * Text wrapping.
     *
     * @var string
     */
    private $wrap = self::WRAP_SQUARE;

    /**
     * Autofit.
     *
     * @var string
     */
    private $autoFit = self::AUTOFIT_DEFAULT;

    /**
     * Horizontal overflow.
     *
     * @var string
     */
    private $horizontalOverflow = self::OVERFLOW_OVERFLOW;

    /**
     * Vertical overflow.
     *
     * @var string
     */
    private $verticalOverflow = self::OVERFLOW_OVERFLOW;

    /**
     * Text upright?
     *
     * @var bool
     */
    private $upright = false;

    /**
     * Vertical text?
     *
     * @var bool
     */
    private $vertical = false;

    /**
     * Number of columns (1 - 16).
     *
     * @var int
     */
    private $columns = 1;

    /**
     * The spacing between columns
     *
     * @var int
     */
    private $columnSpacing = 0;

    /**
     * Bottom inset (in pixels).
     *
     * @var float
     */
    private $bottomInset = 4.8;

    /**
     * Left inset (in pixels).
     *
     * @var float
     */
    private $leftInset = 9.6;

    /**
     * Right inset (in pixels).
     *
     * @var float
     */
    private $rightInset = 9.6;

    /**
     * Top inset (in pixels).
     *
     * @var float
     */
    private $topInset = 4.8;

    /**
     * Horizontal Auto Shrink.
     *
     * @var bool|null
     */
    private $autoShrinkHorizontal;

    /**
     * Vertical Auto Shrink.
     *
     * @var bool|null
     */
    private $autoShrinkVertical;

    /**
     * The percentage of the original font size to which the text is scaled.
     *
     * @var float|null
     */
    private $fontScale;

    /**
     * The percentage of the reduction of the line spacing.
     *
     * @var float|null
     */
    private $lnSpcReduction;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\RichText instance.
     */
    public function __construct()
    {
        // Initialise variables
        $this->richTextParagraphs = [
            new Paragraph(),
        ];

        // Initialize parent
        parent::__construct();
    }

    /**
     * Get active paragraph index.
     *
     * @return int
     */
    public function getActiveParagraphIndex(): int
    {
        return $this->activeParagraph;
    }

    public function getActiveParagraph(): Paragraph
    {
        return $this->richTextParagraphs[$this->activeParagraph];
    }

    /**
     * Set active paragraph.
     *
     * @throws OutOfBoundsException
     */
    public function setActiveParagraph(int $index = 0): Paragraph
    {
        if ($index >= count($this->richTextParagraphs)) {
            throw new OutOfBoundsException(0, count($this->richTextParagraphs), $index);
        }

        $this->activeParagraph = $index;

        return $this->getActiveParagraph();
    }

    /**
     * Get paragraph.
     *
     * @throws OutOfBoundsException
     */
    public function getParagraph(int $index = 0): Paragraph
    {
        if ($index >= count($this->richTextParagraphs)) {
            throw new OutOfBoundsException(0, count($this->richTextParagraphs), $index);
        }

        return $this->richTextParagraphs[$index];
    }

    /**
     * Create paragraph.
     */
    public function createParagraph(): Paragraph
    {
        $numParagraphs = count($this->richTextParagraphs);
        if ($numParagraphs > 0) {
            $alignment = clone $this->getActiveParagraph()->getAlignment();
            $font = clone $this->getActiveParagraph()->getFont();
            $bulletStyle = clone $this->getActiveParagraph()->getBulletStyle();
        }

        $this->richTextParagraphs[] = new Paragraph();
        $this->activeParagraph = count($this->richTextParagraphs) - 1;

        if (isset($alignment)) {
            $this->getActiveParagraph()->setAlignment($alignment);
        }
        if (isset($font)) {
            $this->getActiveParagraph()->setFont($font);
        }
        if (isset($bulletStyle)) {
            $this->getActiveParagraph()->setBulletStyle($bulletStyle);
        }

        return $this->getActiveParagraph();
    }

    /**
     * Add text.
     *
     * @param TextElementInterface|null $pText Rich text element
     *
     * @return self
     */
    public function addText(TextElementInterface $pText = null): self
    {
        $this->richTextParagraphs[$this->activeParagraph]->addText($pText);

        return $this;
    }

    /**
     * Create text (can not be formatted !).
     *
     * @param string $pText Text
     *
     * @return RichText\TextElement
     */
    public function createText(string $pText = ''): RichText\TextElement
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createText($pText);
    }

    /**
     * Create break.
     *
     * @return RichText\BreakElement
     */
    public function createBreak(): RichText\BreakElement
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createBreak();
    }

    /**
     * Create text run (can be formatted).
     *
     * @param string $pText Text
     *
     * @return RichText\Run
     */
    public function createTextRun(string $pText = ''): RichText\Run
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createTextRun($pText);
    }

    /**
     * Get plain text.
     *
     * @return string
     */
    public function getPlainText(): string
    {
        // Return value
        $returnValue = '';

        // Loop trough all Paragraph
        foreach ($this->richTextParagraphs as $p) {
            $returnValue .= $p->getPlainText();
        }

        // Return
        return $returnValue;
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
     * @return array<Paragraph>
     */
    public function getParagraphs(): array
    {
        return $this->richTextParagraphs;
    }

    /**
     * Set paragraphs.
     *
     * @param array<Paragraph> $paragraphs Array of paragraphs
     */
    public function setParagraphs(array $paragraphs = []): self
    {
        $this->richTextParagraphs = $paragraphs;
        $this->activeParagraph = count($this->richTextParagraphs) - 1;

        return $this;
    }

    /**
     * Get text wrapping.
     *
     * @return string
     */
    public function getWrap(): string
    {
        return $this->wrap;
    }

    /**
     * Set text wrapping.
     *
     * @param string $value
     *
     * @return self
     */
    public function setWrap(string $value = self::WRAP_SQUARE): self
    {
        $this->wrap = $value;

        return $this;
    }

    /**
     * Get autofit.
     *
     * @return string
     */
    public function getAutoFit(): string
    {
        return $this->autoFit;
    }

    /**
     * Get pourcentage of fontScale.
     */
    public function getFontScale(): ?float
    {
        return $this->fontScale;
    }

    /**
     * Get pourcentage of the line space reduction.
     */
    public function getLineSpaceReduction(): ?float
    {
        return $this->lnSpcReduction;
    }

    /**
     * Set autofit.
     *
     * @param string $value
     * @param float|null $fontScale
     * @param float|null $lnSpcReduction
     *
     * @return self
     */
    public function setAutoFit(string $value = self::AUTOFIT_DEFAULT, float $fontScale = null, float $lnSpcReduction = null): self
    {
        $this->autoFit = $value;

        if (!is_null($fontScale)) {
            $this->fontScale = $fontScale;
        }

        if (!is_null($lnSpcReduction)) {
            $this->lnSpcReduction = $lnSpcReduction;
        }

        return $this;
    }

    /**
     * Get horizontal overflow.
     *
     * @return string
     */
    public function getHorizontalOverflow(): string
    {
        return $this->horizontalOverflow;
    }

    /**
     * Set horizontal overflow.
     *
     * @param string $value
     *
     * @return self
     */
    public function setHorizontalOverflow(string $value = self::OVERFLOW_OVERFLOW): self
    {
        $this->horizontalOverflow = $value;

        return $this;
    }

    /**
     * Get vertical overflow.
     *
     * @return string
     */
    public function getVerticalOverflow(): string
    {
        return $this->verticalOverflow;
    }

    /**
     * Set vertical overflow.
     *
     * @param string $value
     *
     * @return self
     */
    public function setVerticalOverflow(string $value = self::OVERFLOW_OVERFLOW): self
    {
        $this->verticalOverflow = $value;

        return $this;
    }

    /**
     * Get upright.
     *
     * @return bool
     */
    public function isUpright(): bool
    {
        return $this->upright;
    }

    /**
     * Set vertical.
     *
     * @param bool $value
     *
     * @return self
     */
    public function setUpright(bool $value = false): self
    {
        $this->upright = $value;

        return $this;
    }

    /**
     * Get vertical.
     *
     * @return bool
     */
    public function isVertical(): bool
    {
        return $this->vertical;
    }

    /**
     * Set vertical.
     *
     * @param bool $value
     *
     * @return self
     */
    public function setVertical(bool $value = false): self
    {
        $this->vertical = $value;

        return $this;
    }

    /**
     * Get columns.
     *
     * @return int
     */
    public function getColumns(): int
    {
        return $this->columns;
    }

    /**
     * Set columns.
     *
     * @param int $value
     *
     * @return self
     *
     * @throws OutOfBoundsException
     */
    public function setColumns(int $value = 1): self
    {
        if ($value > 16 || $value < 1) {
            throw new OutOfBoundsException(1, 16, $value);
        }

        $this->columns = $value;

        return $this;
    }

    /**
     * Get bottom inset.
     *
     * @return float
     */
    public function getInsetBottom(): float
    {
        return $this->bottomInset;
    }

    /**
     * Set bottom inset.
     *
     * @param float $value
     *
     * @return self
     */
    public function setInsetBottom(float $value = 4.8): self
    {
        $this->bottomInset = $value;

        return $this;
    }

    /**
     * Get left inset.
     *
     * @return float
     */
    public function getInsetLeft(): float
    {
        return $this->leftInset;
    }

    /**
     * Set left inset.
     *
     * @param float $value
     *
     * @return self
     */
    public function setInsetLeft(float $value = 9.6): self
    {
        $this->leftInset = $value;

        return $this;
    }

    /**
     * Get right inset.
     *
     * @return float
     */
    public function getInsetRight(): float
    {
        return $this->rightInset;
    }

    /**
     * Set left inset.
     *
     * @param float $value
     *
     * @return self
     */
    public function setInsetRight(float $value = 9.6): self
    {
        $this->rightInset = $value;

        return $this;
    }

    /**
     * Get top inset.
     *
     * @return float
     */
    public function getInsetTop(): float
    {
        return $this->topInset;
    }

    /**
     * Set top inset.
     *
     * @param float $value
     *
     * @return self
     */
    public function setInsetTop(float $value = 4.8): self
    {
        $this->topInset = $value;

        return $this;
    }

    public function setAutoShrinkHorizontal(bool $value = null): self
    {
        $this->autoShrinkHorizontal = $value;

        return $this;
    }

    public function hasAutoShrinkHorizontal(): ?bool
    {
        return $this->autoShrinkHorizontal;
    }

    /**
     * Set vertical auto shrink.
     *
     * @return RichText
     */
    public function setAutoShrinkVertical(bool $value = null): self
    {
        $this->autoShrinkVertical = $value;

        return $this;
    }

    /**
     * Set vertical auto shrink.
     */
    public function hasAutoShrinkVertical(): ?bool
    {
        return $this->autoShrinkVertical;
    }

    /**
     * Get spacing between columns
     *
     * @return int
     */
    public function getColumnSpacing(): int
    {
        return $this->columnSpacing;
    }

    /**
     * Set spacing between columns
     *
     * @param int $value
     *
     * @return self
     */
    public function setColumnSpacing(int $value = 0): self
    {
        if ($value >= 0) {
            $this->columnSpacing = $value;
        }

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
        foreach ($this->richTextParagraphs as $element) {
            $hashElements .= $element->getHashCode();
        }

        return md5(
            $hashElements
            . $this->wrap
            . $this->autoFit
            . $this->horizontalOverflow
            . $this->verticalOverflow
            . ($this->upright ? '1' : '0')
            . ($this->vertical ? '1' : '0')
            . $this->columns
            . $this->columnSpacing
            . $this->bottomInset
            . $this->leftInset
            . $this->rightInset
            . $this->topInset
            . parent::getHashCode()
            . __CLASS__
        );
    }
}
