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

namespace PhpOffice\PhpPresentation\Shape\Table;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface;
use PhpOffice\PhpPresentation\Style\Borders;
use PhpOffice\PhpPresentation\Style\Fill;

/**
 * Table cell.
 */
class Cell implements ComparableInterface
{
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
     * Fill.
     *
     * @var Fill
     */
    private $fill;

    /**
     * Borders.
     *
     * @var Borders
     */
    private $borders;

    /**
     * Width (in pixels).
     *
     * @var int
     */
    private $width = 0;

    /**
     * Colspan.
     *
     * @var int
     */
    private $colSpan = 0;

    /**
     * Rowspan.
     *
     * @var int
     */
    private $rowSpan = 0;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\RichText instance.
     */
    public function __construct()
    {
        // Initialise variables
        $this->richTextParagraphs = [
            new Paragraph(),
        ];
        $this->activeParagraph = 0;

        // Set fill
        $this->fill = new Fill();

        // Set borders
        $this->borders = new Borders();
    }

    /**
     * Get active paragraph index.
     *
     * @return int
     */
    public function getActiveParagraphIndex()
    {
        return $this->activeParagraph;
    }

    /**
     * Get active paragraph.
     */
    public function getActiveParagraph(): Paragraph
    {
        return $this->richTextParagraphs[$this->activeParagraph];
    }

    /**
     * Set active paragraph.
     *
     * @param int $index
     */
    public function setActiveParagraph($index = 0): Paragraph
    {
        if ($index >= count($this->richTextParagraphs)) {
            throw new OutOfBoundsException(0, count($this->richTextParagraphs), $index);
        }

        $this->activeParagraph = $index;

        return $this->getActiveParagraph();
    }

    /**
     * Get paragraph.
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
        $this->richTextParagraphs[] = new Paragraph();
        $totalRichTextParagraphs = count($this->richTextParagraphs);
        $this->activeParagraph = $totalRichTextParagraphs - 1;

        if ($totalRichTextParagraphs > 1) {
            $alignment = clone $this->getActiveParagraph()->getAlignment();
            $font = clone $this->getActiveParagraph()->getFont();
            $bulletStyle = clone $this->getActiveParagraph()->getBulletStyle();

            $this->getActiveParagraph()->setAlignment($alignment);
            $this->getActiveParagraph()->setFont($font);
            $this->getActiveParagraph()->setBulletStyle($bulletStyle);
        }

        return $this->getActiveParagraph();
    }

    /**
     * Add text.
     *
     * @param TextElementInterface $pText Rich text element
     *
     * @return Cell
     */
    public function addText(?TextElementInterface $pText = null)
    {
        $this->richTextParagraphs[$this->activeParagraph]->addText($pText);

        return $this;
    }

    /**
     * Create text (can not be formatted !).
     *
     * @param string $pText Text
     *
     * @return \PhpOffice\PhpPresentation\Shape\RichText\TextElement
     */
    public function createText($pText = '')
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createText($pText);
    }

    /**
     * Create break.
     *
     * @return \PhpOffice\PhpPresentation\Shape\RichText\BreakElement
     */
    public function createBreak()
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createBreak();
    }

    /**
     * Create text run (can be formatted).
     *
     * @param string $pText Text
     *
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Run
     */
    public function createTextRun(string $pText = '')
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createTextRun($pText);
    }

    /**
     * Get plain text.
     *
     * @return string
     */
    public function getPlainText()
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
     * Get paragraphs.
     *
     * @return array<Paragraph>
     */
    public function getParagraphs()
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
     * Get fill.
     *
     * @return Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Set fill.
     *
     * @return Cell
     */
    public function setFill(Fill $fill)
    {
        $this->fill = $fill;

        return $this;
    }

    /**
     * Get borders.
     *
     * @return Borders
     */
    public function getBorders()
    {
        return $this->borders;
    }

    /**
     * Set borders.
     *
     * @return Cell
     */
    public function setBorders(Borders $borders)
    {
        $this->borders = $borders;

        return $this;
    }

    /**
     * Get width.
     *
     * @return int
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set width.
     *
     * @return self
     */
    public function setWidth(int $pValue = 0)
    {
        $this->width = $pValue;

        return $this;
    }

    public function getColSpan(): int
    {
        return $this->colSpan;
    }

    public function setColSpan(int $value = 0): self
    {
        $this->colSpan = $value;

        return $this;
    }

    public function getRowSpan(): int
    {
        return $this->rowSpan;
    }

    public function setRowSpan(int $value = 0): self
    {
        $this->rowSpan = $value;

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

        return md5($hashElements . $this->fill->getHashCode() . $this->borders->getHashCode() . $this->width . __CLASS__);
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
}
