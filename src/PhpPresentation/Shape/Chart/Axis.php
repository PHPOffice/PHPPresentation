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
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Axis
 */
class Axis implements ComparableInterface
{
    const AXIS_X = 'x';
    const AXIS_Y = 'y';

    const TICK_MARK_NONE = 'none';
    const TICK_MARK_CROSS = 'cross';
    const TICK_MARK_INSIDE = 'in';
    const TICK_MARK_OUTSIDE = 'out';

    /**
     * Title
     *
     * @var string
     */
    private $title = 'Axis Title';

    /**
     * Format code
     *
     * @var string
     */
    private $formatCode = '';

    /**
     * Font
     *
     * @var \PhpOffice\PhpPresentation\Style\Font
     */
    private $font;

    /**
     * @var Gridlines
     */
    protected $majorGridlines;

    /**
     * @var Gridlines
     */
    protected $minorGridlines;

    /**
     * @var int
     */
    protected $minBounds;

    /**
     * @var int
     */
    protected $maxBounds;

    /**
     * @var string
     */
    protected $minorTickMark = self::TICK_MARK_NONE;

    /**
     * @var string
     */
    protected $majorTickMark = self::TICK_MARK_NONE;

    /**
     * @var float
     */
    protected $minorUnit;

    /**
     * @var float
     */
    protected $majorUnit;

    /**
     * @var Outline
     */
    protected $outline;

    /**
     * @var boolean
     */
    protected $isVisible = true;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Chart\Axis instance
     *
     * @param string $title Title
     */
    public function __construct($title = 'Axis Title')
    {
        $this->title = $title;
        $this->outline = new Outline();
        $this->font  = new Font();
    }

    /**
     * Get Title
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->title;
    }

    /**
     * Set Title
     *
     * @param  string                         $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    public function setTitle($value = 'Axis Title')
    {
        $this->title = $value;

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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    public function setFont(Font $pFont = null)
    {
        $this->font = $pFont;
        return $this;
    }

    /**
     * Get Format Code
     *
     * @return string
     */
    public function getFormatCode()
    {
        return $this->formatCode;
    }

    /**
     * Set Format Code
     *
     * @param  string                         $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    public function setFormatCode($value = '')
    {
        $this->formatCode = $value;

        return $this;
    }

    /**
     * @return int|null
     */
    public function getMinBounds()
    {
        return $this->minBounds;
    }

    /**
     * @param int|null $minBounds
     * @return Axis
     */
    public function setMinBounds($minBounds = null)
    {
        $this->minBounds = is_null($minBounds) ? null : (int)$minBounds;
        return $this;
    }

    /**
     * @return int|null
     */
    public function getMaxBounds()
    {
        return $this->maxBounds;
    }

    /**
     * @param int|null $maxBounds
     * @return Axis
     */
    public function setMaxBounds($maxBounds = null)
    {
        $this->maxBounds = is_null($maxBounds) ? null : (int)$maxBounds;
        return $this;
    }

    /**
     * @return Gridlines
     */
    public function getMajorGridlines()
    {
        return $this->majorGridlines;
    }

    /**
     * @param Gridlines $majorGridlines
     * @return Axis
     */
    public function setMajorGridlines(Gridlines $majorGridlines)
    {
        $this->majorGridlines = $majorGridlines;
        return $this;
    }

    /**
     * @return Gridlines
     */
    public function getMinorGridlines()
    {
        return $this->minorGridlines;
    }

    /**
     * @param Gridlines $minorGridlines
     * @return Axis
     */
    public function setMinorGridlines(Gridlines $minorGridlines)
    {
        $this->minorGridlines = $minorGridlines;
        return $this;
    }

    /**
     * @return string
     */
    public function getMinorTickMark()
    {
        return $this->minorTickMark;
    }

    /**
     * @param string $pTickMark
     * @return Axis
     */
    public function setMinorTickMark($pTickMark = self::TICK_MARK_NONE)
    {
        $this->minorTickMark = $pTickMark;
        return $this;
    }

    /**
     * @return string
     */
    public function getMajorTickMark()
    {
        return $this->majorTickMark;
    }

    /**
     * @param string $pTickMark
     * @return Axis
     */
    public function setMajorTickMark($pTickMark = self::TICK_MARK_NONE)
    {
        $this->majorTickMark = $pTickMark;
        return $this;
    }

    /**
     * @return float
     */
    public function getMinorUnit()
    {
        return $this->minorUnit;
    }

    /**
     * @param float $pUnit
     * @return Axis
     */
    public function setMinorUnit($pUnit = null)
    {
        $this->minorUnit = $pUnit;
        return $this;
    }

    /**
     * @return float
     */
    public function getMajorUnit()
    {
        return $this->majorUnit;
    }

    /**
     * @param float $pUnit
     * @return Axis
     */
    public function setMajorUnit($pUnit = null)
    {
        $this->majorUnit = $pUnit;
        return $this;
    }

    /**
     * @return Outline
     */
    public function getOutline()
    {
        return $this->outline;
    }

    /**
     * @param Outline $outline
     * @return Axis
     */
    public function setOutline(Outline $outline)
    {
        $this->outline = $outline;
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->title . $this->formatCode . __CLASS__);
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
     * @return $this
     */
    public function setHashIndex($value)
    {
        $this->hashIndex = $value;
        return $this;
    }

    /**
     * Axis is hidden ?
     * @return boolean
     */
    public function isVisible()
    {
        return $this->isVisible;
    }

    /**
     * Hide an axis
     *
     * @param boolean $value delete
     * @return $this
     */
    public function setIsVisible($value)
    {
        $this->isVisible = (bool)$value;
        return $this;
    }
}
