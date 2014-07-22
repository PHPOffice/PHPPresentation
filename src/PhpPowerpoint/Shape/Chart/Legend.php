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

namespace PhpOffice\PhpPowerpoint\Shape\Chart;

use PhpOffice\PhpPowerpoint\ComparableInterface;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Border;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Style\Font;

/**
 * \PhpOffice\PhpPowerpoint\Shape\Chart\Legend
 */
class Legend implements ComparableInterface
{
    /** Legend positions */
    const POSITION_BOTTOM = 'b';
    const POSITION_LEFT = 'l';
    const POSITION_RIGHT = 'r';
    const POSITION_TOP = 't';
    const POSITION_TOPRIGHT = 'tr';

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
     * @var \PhpOffice\PhpPowerpoint\Style\Font
     */
    private $font;

    /**
     * Border
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Border
     */
    private $border;

    /**
     * Fill
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Fill
     */
    private $fill;

    /**
     * Alignment
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Alignment
     */
    private $alignment;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Shape\Chart\Legend instance
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
     * @param  boolean                          $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Legend
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
     * @param  string                          $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Title
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
    public function getOffsetX()
    {
        return $this->offsetX;
    }

    /**
     * Set OffsetX (as a fraction of the chart)
     *
     * @param float|int $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Legend
     */
    public function setOffsetX($value = 0)
    {
        $this->offsetX = (double)$value;
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
     * @param float|int $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Legend
     */
    public function setOffsetY($value = 0)
    {
        $this->offsetY = (double)$value;
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Legend
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Legend
     */
    public function setHeight($value = 0)
    {
        $this->height = (double)$value;
        return $this;
    }

    /**
     * Get font
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param  \PhpOffice\PhpPowerpoint\Style\Font               $pFont Font
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph
     */
    public function setFont(Font $pFont = null)
    {
        $this->font = $pFont;
        return $this;
    }

    /**
     * Get Border
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Border
     */
    public function getBorder()
    {
        return $this->border;
    }

    /**
     * Get Fill
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Get alignment
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Alignment
     */
    public function getAlignment()
    {
        return $this->alignment;
    }

    /**
     * Set alignment
     *
     * @param  \PhpOffice\PhpPowerpoint\Style\Alignment          $alignment
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph
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
        return md5($this->position . $this->offsetX . $this->offsetY . $this->width . $this->height . $this->font->getHashCode() . $this->border->getHashCode() . $this->fill->getHashCode() . $this->alignment->getHashCode() . ($this->visible ? 't' : 'f') . __CLASS__);
    }

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

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
