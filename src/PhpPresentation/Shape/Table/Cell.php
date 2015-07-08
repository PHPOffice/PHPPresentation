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

namespace PhpOffice\PhpPresentation\Shape\Table;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface;
use PhpOffice\PhpPresentation\Style\Borders;
use PhpOffice\PhpPresentation\Style\Fill;

/**
 * Table cell
 */
class Cell implements ComparableInterface
{
    /**
     * Rich text paragraphs
     *
     * @var \PhpOffice\PhpPresentation\Shape\RichText\Paragraph[]
     */
    private $richTextParagraphs;

    /**
     * Active paragraph
     *
     * @var int
     */
    private $activeParagraph = 0;

    /**
     * Fill
     *
     * @var \PhpOffice\PhpPresentation\Style\Fill
     */
    private $fill;

    /**
     * Borders
     *
     * @var \PhpOffice\PhpPresentation\Style\Borders
     */
    private $borders;

    /**
     * Width (in pixels)
     *
     * @var int
     */
    private $width = 0;

    /**
     * Colspan
     *
     * @var int
     */
    private $colSpan = 0;

    /**
     * Rowspan
     *
     * @var int
     */
    private $rowSpan = 0;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\RichText instance
     */
    public function __construct()
    {
        // Initialise variables
        $this->richTextParagraphs = array(
            new Paragraph()
        );
        $this->activeParagraph    = 0;

        // Set fill
        $this->fill = new Fill();

        // Set borders
        $this->borders = new Borders();
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
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
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
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
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
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
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
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
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
     * @param  \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface $pText Rich text element
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\RichText
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
     * @return \PhpOffice\PhpPresentation\Shape\RichText\TextElement
     * @throws \Exception
     */
    public function createText($pText = '')
    {
        return $this->richTextParagraphs[$this->activeParagraph]->createText($pText);
    }

    /**
     * Create break
     *
     * @return \PhpOffice\PhpPresentation\Shape\RichText\BreakElement
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
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Run
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

        // Loop trough all \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
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
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph[]
     */
    public function getParagraphs()
    {
        return $this->richTextParagraphs;
    }

    /**
     * Set paragraphs
     *
     * @param  \PhpOffice\PhpPresentation\Shape\RichText\Paragraph[] $paragraphs Array of paragraphs
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\RichText
     */
    public function setParagraphs($paragraphs = null)
    {
        if (is_array($paragraphs)) {
            $this->richTextParagraphs = $paragraphs;
            $this->activeParagraph    = count($this->richTextParagraphs) - 1;
        } else {
            throw new \Exception("Invalid \PhpOffice\PhpPresentation\Shape\RichText\Paragraph[] array passed.");
        }

        return $this;
    }

    /**
     * Get fill
     *
     * @return \PhpOffice\PhpPresentation\Style\Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Set fill
     *
     * @param  \PhpOffice\PhpPresentation\Style\Fill     $fill
     * @return \PhpOffice\PhpPresentation\Shape\RichText
     */
    public function setFill(Fill $fill)
    {
        $this->fill = $fill;

        return $this;
    }

    /**
     * Get borders
     *
     * @return \PhpOffice\PhpPresentation\Style\Borders
     */
    public function getBorders()
    {
        return $this->borders;
    }

    /**
     * Set borders
     *
     * @param  \PhpOffice\PhpPresentation\Style\Borders  $borders
     * @return \PhpOffice\PhpPresentation\Shape\RichText
     */
    public function setBorders(Borders $borders)
    {
        $this->borders = $borders;

        return $this;
    }

    /**
     * Get width
     *
     * @return int
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set width
     *
     * @param  int                          $value
     * @return \PhpOffice\PhpPresentation\Shape\RichText
     */
    public function setWidth($value = 0)
    {
        $this->width = $value;

        return $this;
    }

    /**
     * Get colSpan
     *
     * @return int
     */
    public function getColSpan()
    {
        return $this->colSpan;
    }

    /**
     * Set colSpan
     *
     * @param  int                          $value
     * @return \PhpOffice\PhpPresentation\Shape\RichText
     */
    public function setColSpan($value = 0)
    {
        $this->colSpan = $value;

        return $this;
    }

    /**
     * Get rowSpan
     *
     * @return int
     */
    public function getRowSpan()
    {
        return $this->rowSpan;
    }

    /**
     * Set rowSpan
     *
     * @param  int                          $value
     * @return \PhpOffice\PhpPresentation\Shape\RichText
     */
    public function setRowSpan($value = 0)
    {
        $this->rowSpan = $value;

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
        foreach ($this->richTextParagraphs as $element) {
            $hashElements .= $element->getHashCode();
        }

        return md5($hashElements . $this->fill->getHashCode() . $this->borders->getHashCode() . $this->width . __CLASS__);
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
