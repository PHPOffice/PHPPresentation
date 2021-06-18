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
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Legend
 */
class Legend implements ComparableInterface
{
    /** Legend positions */
    public const POSITION_BOTTOM = 'b';
    public const POSITION_LEFT = 'l';
    public const POSITION_RIGHT = 'r';
    public const POSITION_TOP = 't';
    public const POSITION_TOPRIGHT = 'tr';

    /**
     * Visible
     *
     * @var boolean
     */
    private $visible = true;

    /**
     * Position
     *
     * @var string
     */
    private $position = self::POSITION_RIGHT;

    /**
     * OffsetX (as a fraction of the chart)
     *
     * @var float
     */
    private $offsetX = 0;

    /**
     * OffsetY (as a fraction of the chart)
     *
     * @var float
     */
    private $offsetY = 0;

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
     * Font
     *
     * @var Font|null
     */
    private $font;

    /**
     * Border
     *
     * @var \PhpOffice\PhpPresentation\Style\Border
     */
    private $border;

    /**
     * Fill
     *
     * @var \PhpOffice\PhpPresentation\Style\Fill
     */
    private $fill;

    /**
     * Alignment
     *
     * @var \PhpOffice\PhpPresentation\Style\Alignment
     */
    private $alignment;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Chart\Legend instance
     */
    public function __construct()
    {
        $this->font      = new Font();
        $this->border    = new Border();
        $this->fill      = new Fill();
        $this->alignment = new Alignment();
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
     * @param boolean                          $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Legend
     */
    public function setVisible($value = true)
    {
        $this->visible = $value;
        return $this;
    }

    /**
     * Get Position
     *
     * @return string
     */
    public function getPosition()
    {
        return $this->position;
    }

    /**
     * Set Position
     *
     * @param string                          $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Legend
     */
    public function setPosition($value = self::POSITION_RIGHT)
    {
        $this->position = $value;
        return $this;
    }

    /**
     * Get OffsetX (as a fraction of the chart)
     *
     * @return float
     */
    public function getOffsetX(): float
    {
        return $this->offsetX;
    }

    /**
     * Set OffsetX (as a fraction of the chart)
     *
     * @param float $pValue
     * @return self
     */
    public function setOffsetX(float $pValue = 0): self
    {
        $this->offsetX = $pValue;
        return $this;
    }

    /**
     * Get OffsetY (as a fraction of the chart)
     *
     * @return float
     */
    public function getOffsetY(): float
    {
        return $this->offsetY;
    }

    /**
     * Set OffsetY (as a fraction of the chart)
     *
     * @param float $pValue
     * @return self
     */
    public function setOffsetY(float $pValue = 0): self
    {
        $this->offsetY = $pValue;
        return $this;
    }

    /**
     * Get Width (as a fraction of the chart)
     *
     * @return float
     */
    public function getWidth(): float
    {
        return $this->width;
    }

    /**
     * Set Width (as a fraction of the chart)
     *
     * @param float $pValue
     * @return self
     */
    public function setWidth(float $pValue = 0): self
    {
        $this->width = $pValue;
        return $this;
    }

    /**
     * Get Height (as a fraction of the chart)
     *
     * @return float
     */
    public function getHeight(): float
    {
        return $this->height;
    }

    /**
     * Set Height (as a fraction of the chart)
     *
     * @param float $value
     * @return self
     */
    public function setHeight(float $value = 0): self
    {
        $this->height = $value;
        return $this;
    }

    /**
     * Get font
     *
     * @return Font|null
     */
    public function getFont(): ?Font
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param Font|null $pFont Font
     * @throws \Exception
     * @return self
     */
    public function setFont(Font $pFont = null): self
    {
        $this->font = $pFont;
        return $this;
    }

    /**
     * Get Border
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function getBorder()
    {
        return $this->border;
    }

    /**
     * Set Border
     *
     * @param \PhpOffice\PhpPresentation\Style\Border $border
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Legend
     */
    public function setBorder(Border $border)
    {
        $this->border = $border;
        return $this;
    }

    /**
     * Get Fill
     *
     * @return \PhpOffice\PhpPresentation\Style\Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Set Fill
     *
     * @param \PhpOffice\PhpPresentation\Style\Fill $fill
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Legend
     */
    public function setFill(Fill $fill)
    {
        $this->fill = $fill;
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
     * @param \PhpOffice\PhpPresentation\Style\Alignment          $alignment
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Legend
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
    public function getHashCode(): string
    {
        return md5($this->position . $this->offsetX . $this->offsetY . $this->width . $this->height . $this->font->getHashCode() . $this->border->getHashCode() . $this->fill->getHashCode() . $this->alignment->getHashCode() . ($this->visible ? 't' : 'f') . __CLASS__);
    }

    /**
     * Hash index
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Get hash index
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
     * Set hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param int $value Hash index
     * @return Legend
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;
        return $this;
    }
}
