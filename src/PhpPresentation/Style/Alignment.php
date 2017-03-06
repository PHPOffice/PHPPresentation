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

namespace PhpOffice\PhpPresentation\Style;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\Style\Alignment
 */
class Alignment implements ComparableInterface
{
    /* Horizontal alignment */
    const HORIZONTAL_GENERAL                = 'l';
    const HORIZONTAL_LEFT                   = 'l';
    const HORIZONTAL_RIGHT                  = 'r';
    const HORIZONTAL_CENTER                 = 'ctr';
    const HORIZONTAL_JUSTIFY                = 'just';
    const HORIZONTAL_DISTRIBUTED            = 'dist';

    /* Vertical alignment */
    const VERTICAL_BASE                     = 'base';
    const VERTICAL_AUTO                     = 'auto';
    const VERTICAL_BOTTOM                   = 'b';
    const VERTICAL_TOP                      = 't';
    const VERTICAL_CENTER                   = 'ctr';

    /* Text direction */
    const TEXT_DIRECTION_HORIZONTAL = 'horz';
    const TEXT_DIRECTION_VERTICAL_90 = 'vert';
    const TEXT_DIRECTION_VERTICAL_270 = 'vert270';
    const TEXT_DIRECTION_STACKED = 'wordArtVert';

    private $supportedStyles = array(
        self::HORIZONTAL_GENERAL,
        self::HORIZONTAL_LEFT,
        self::HORIZONTAL_RIGHT,
    );

    /**
     * Horizontal
     * @var string
     */
    private $horizontal;

    /**
     * Vertical
     * @var string
     */
    private $vertical;

    /**
     * Text Direction
     * @var string
     */
    private $textDirection = self::TEXT_DIRECTION_HORIZONTAL;

    /**
     * Level
     * @var int
     */
    private $level = 0;

    /**
     * Indent - only possible with horizontal alignment left and right
     * @var int
     */
    private $indent = 0;

    /**
     * Margin left - only possible with horizontal alignment left and right
     * @var int
     */
    private $marginLeft = 0;

    /**
     * Margin right - only possible with horizontal alignment left and right
     * @var int
     */
    private $marginRight = 0;

    /**
     * Margin top
     * @var int
     */
    private $marginTop = 0;

    /**
     * Margin bottom
     * @var int
     */
    private $marginBottom = 0;

    /**
     * Hash index
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function __construct()
    {
        // Initialise values
        $this->horizontal          = self::HORIZONTAL_LEFT;
        $this->vertical            = self::VERTICAL_BASE;
    }

    /**
     * Get Horizontal
     *
     * @return string
     */
    public function getHorizontal()
    {
        return $this->horizontal;
    }

    /**
     * Set Horizontal
     *
     * @param  string                        $pValue
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function setHorizontal($pValue = self::HORIZONTAL_LEFT)
    {
        if ($pValue == '') {
            $pValue = self::HORIZONTAL_LEFT;
        }
        $this->horizontal = $pValue;

        return $this;
    }

    /**
     * Get Vertical
     *
     * @return string
     */
    public function getVertical()
    {
        return $this->vertical;
    }

    /**
     * Set Vertical
     *
     * @param  string                        $pValue
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function setVertical($pValue = self::VERTICAL_BASE)
    {
        if ($pValue == '') {
            $pValue = self::VERTICAL_BASE;
        }
        $this->vertical = $pValue;

        return $this;
    }

    /**
     * Get Level
     *
     * @return int
     */
    public function getLevel()
    {
        return $this->level;
    }

    /**
     * Set Level
     *
     * @param  int                           $pValue Ranging 0 - 8
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function setLevel($pValue = 0)
    {
        if ($pValue < 0) {
            throw new \Exception("Invalid value should be more than 0.");
        }
        $this->level = $pValue;

        return $this;
    }

    /**
     * Get indent
     *
     * @return int
     */
    public function getIndent()
    {
        return $this->indent;
    }

    /**
     * Set indent
     *
     * @param  int                           $pValue
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function setIndent($pValue = 0)
    {
        if ($pValue > 0 && !in_array($this->getHorizontal(), $this->supportedStyles)) {
            $pValue = 0; // indent not supported
        }

        $this->indent = $pValue;

        return $this;
    }

    /**
     * Get margin left
     *
     * @return int
     */
    public function getMarginLeft()
    {
        return $this->marginLeft;
    }

    /**
     * Set margin left
     *
     * @param  int                           $pValue
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function setMarginLeft($pValue = 0)
    {
        if ($pValue > 0 && !in_array($this->getHorizontal(), $this->supportedStyles)) {
            $pValue = 0; // margin left not supported
        }

        $this->marginLeft = $pValue;

        return $this;
    }

    /**
     * Get margin right
     *
     * @return int
     */
    public function getMarginRight()
    {
        return $this->marginRight;
    }

    /**
     * Set margin ight
     *
     * @param  int                           $pValue
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function setMarginRight($pValue = 0)
    {
        if ($pValue > 0 && !in_array($this->getHorizontal(), $this->supportedStyles)) {
            $pValue = 0; // margin right not supported
        }

        $this->marginRight = $pValue;

        return $this;
    }

    /**
     * Get margin top
     *
     * @return int
     */
    public function getMarginTop()
    {
        return $this->marginTop;
    }

    /**
     * Set margin top
     *
     * @param  int                           $pValue
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function setMarginTop($pValue = 0)
    {
        $this->marginTop = $pValue;

        return $this;
    }

    /**
     * Get margin bottom
     *
     * @return int
     */
    public function getMarginBottom()
    {
        return $this->marginBottom;
    }

    /**
     * Set margin bottom
     *
     * @param  int                           $pValue
     * @return \PhpOffice\PhpPresentation\Style\Alignment
     */
    public function setMarginBottom($pValue = 0)
    {
        $this->marginBottom = $pValue;

        return $this;
    }

    /**
     * @return string
     */
    public function getTextDirection()
    {
        return $this->textDirection;
    }

    /**
     * @param string $pValue
     * @return Alignment
     */
    public function setTextDirection($pValue = self::TEXT_DIRECTION_HORIZONTAL)
    {
        if (empty($pValue)) {
            $pValue = self::TEXT_DIRECTION_HORIZONTAL;
        }
        $this->textDirection = $pValue;
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5(
            $this->horizontal
            . $this->vertical
            . $this->level
            . $this->indent
            . $this->marginLeft
            . $this->marginRight
            . __CLASS__
        );
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
