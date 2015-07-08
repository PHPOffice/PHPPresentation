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

namespace PhpOffice\PhpPresentation\Shape\Chart;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Font;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Title
 */
class Title implements ComparableInterface
{
    /**
     * Visible
     *
     * @var boolean
     */
    private $visible = true;

    /**
     * Text
     *
     * @var string
     */
    private $text = 'Chart Title';

    /**
     * OffsetX (as a fraction of the chart)
     *
     * @var float
     */
    private $offsetX = 0.01;

    /**
     * OffsetY (as a fraction of the chart)
     *
     * @var float
     */
    private $offsetY = 0.01;

    /**
     * Width (as a fraction of the chart)
     *
     * @var float
     */
    private $width = 0;

    /**
     * Height (as a fraction of the chart)
     *
     * @var float
     */
    private $height = 0;

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
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Chart\Title instance
     */
    public function __construct()
    {
        $this->alignment = new Alignment();
        $this->font      = new Font();
        $this->font->setName('Calibri');
        $this->font->setSize(18);
    }

    /**
     * Get Visible
     *
     * @return boolean
     */
    public function isVisible()
    {
        return $this->visible;
    }

    /**
     * Set Visible
     *
     * @param  boolean                         $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
     */
    public function setVisible($value = true)
    {
        $this->visible = $value;

        return $this;
    }

    /**
     * Get Text
     *
     * @return string
     */
    public function getText()
    {
        return $this->text;
    }

    /**
     * Set Text
     *
     * @param  string                          $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
     */
    public function setText($value = null)
    {
        $this->text = $value;

        return $this;
    }

    /**
     * Get OffsetX (as a fraction of the chart)
     *
     * @return float
     */
    public function getOffsetX()
    {
        return $this->offsetX;
    }

    /**
     * Set OffsetX (as a fraction of the chart)
     *
     * @param  float                           $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
     */
    public function setOffsetX($value = 0.01)
    {
        $this->offsetX = $value;

        return $this;
    }

    /**
     * Get OffsetY (as a fraction of the chart)
     *
     * @return float
     */
    public function getOffsetY()
    {
        return $this->offsetY;
    }

    /**
     * Set OffsetY (as a fraction of the chart)
     *
     * @param  float                           $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
     */
    public function setOffsetY($value = 0.01)
    {
        $this->offsetY = $value;

        return $this;
    }

    /**
     * Get Width (as a fraction of the chart)
     *
     * @return float
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set Width (as a fraction of the chart)
     *
     * @param float|int $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
     */
    public function setWidth($value = 0)
    {
        $this->width = (double)$value;

        return $this;
    }

    /**
     * Get Height (as a fraction of the chart)
     *
     * @return float
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Set Height (as a fraction of the chart)
     *
     * @param float|int $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
     */
    public function setHeight($value = 0)
    {
        $this->height = (double)$value;

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
     * @param  \PhpOffice\PhpPresentation\Style\Font               $pFont Font
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\RichText\Paragraph
     */
    public function setFont(Font $pFont = null)
    {
        $this->font = $pFont;

        return $this;
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
     * @param  \PhpOffice\PhpPresentation\Style\Alignment   $alignment
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
     */
    public function setAlignment(Alignment $alignment)
    {
        $this->alignment = $alignment;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->text . $this->offsetX . $this->offsetY . $this->width . $this->height . $this->font->getHashCode() . $this->alignment->getHashCode() . ($this->visible ? 't' : 'f') . __CLASS__);
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
        return $this;
    }
}
