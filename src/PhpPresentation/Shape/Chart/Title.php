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

namespace PhpOffice\PhpPresentation\Shape\Chart;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Font;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Title.
 */
class Title implements ComparableInterface
{
    /**
     * Visible.
     *
     * @var bool
     */
    private $visible = true;

    /**
     * Text.
     *
     * @var string
     */
    private $text = 'Chart Title';

    /**
     * OffsetX (as a fraction of the chart).
     *
     * @var float
     */
    private $offsetX = 0.01;

    /**
     * OffsetY (as a fraction of the chart).
     *
     * @var float
     */
    private $offsetY = 0.01;

    /**
     * Width (as a fraction of the chart).
     *
     * @var float
     */
    private $width = 0;

    /**
     * Height (as a fraction of the chart).
     *
     * @var float
     */
    private $height = 0;

    /**
     * Alignment.
     *
     * @var Alignment
     */
    private $alignment;

    /**
     * Font.
     *
     * @var Font
     */
    private $font;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Chart\Title instance.
     */
    public function __construct()
    {
        $this->alignment = new Alignment();
        $this->font = new Font();
        $this->font->setName('Calibri');
        $this->font->setSize(18);
    }

    /**
     * Get Visible.
     *
     * @return bool
     */
    public function isVisible()
    {
        return $this->visible;
    }

    /**
     * Set Visible.
     *
     * @param bool $value
     *
     * @return Title
     */
    public function setVisible($value = true)
    {
        $this->visible = $value;

        return $this;
    }

    /**
     * Get Text.
     *
     * @return string
     */
    public function getText()
    {
        return $this->text;
    }

    /**
     * Set Text.
     *
     * @param string $value
     *
     * @return Title
     */
    public function setText($value = null)
    {
        $this->text = $value;

        return $this;
    }

    /**
     * Get OffsetX (as a fraction of the chart).
     */
    public function getOffsetX(): float
    {
        return $this->offsetX;
    }

    /**
     * Set OffsetX (as a fraction of the chart).
     */
    public function setOffsetX(float $value = 0.01): self
    {
        $this->offsetX = $value;

        return $this;
    }

    /**
     * Get OffsetY (as a fraction of the chart).
     */
    public function getOffsetY(): float
    {
        return $this->offsetY;
    }

    /**
     * Set OffsetY (as a fraction of the chart).
     */
    public function setOffsetY(float $pValue = 0.01): self
    {
        $this->offsetY = $pValue;

        return $this;
    }

    /**
     * Get Width (as a fraction of the chart).
     */
    public function getWidth(): float
    {
        return $this->width;
    }

    /**
     * Set Width (as a fraction of the chart).
     */
    public function setWidth(float $pValue = 0): self
    {
        $this->width = $pValue;

        return $this;
    }

    /**
     * Get Height (as a fraction of the chart).
     */
    public function getHeight(): float
    {
        return $this->height;
    }

    /**
     * Set Height (as a fraction of the chart).
     */
    public function setHeight(float $value = 0): self
    {
        $this->height = $value;

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
     * Get alignment.
     *
     * @return Alignment
     */
    public function getAlignment()
    {
        return $this->alignment;
    }

    /**
     * Set alignment.
     *
     * @return Title
     */
    public function setAlignment(Alignment $alignment)
    {
        $this->alignment = $alignment;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5($this->text . $this->offsetX . $this->offsetY . $this->width . $this->height . $this->font->getHashCode() . $this->alignment->getHashCode() . ($this->visible ? 't' : 'f') . __CLASS__);
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
     * @return Title
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;

        return $this;
    }
}
