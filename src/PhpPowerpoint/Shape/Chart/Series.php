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

use PhpOffice\PhpPowerpoint\IComparable;
use PhpOffice\PhpPowerpoint\Style\Fill;
use PhpOffice\PhpPowerpoint\Style\Font;

/**
 * PHPPowerPoint_Shape_Chart_Series
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape_Chart
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Series implements IComparable
{
    /* Label positions */
    const LABEL_BESTFIT = 'bestFir';
    const LABEL_BOTTOM = 'b';
    const LABEL_CENTER = 'ctr';
    const LABEL_INSIDEBASE = 'inBase';
    const LABEL_INSIDEEND = 'inEnd';
    const LABEL_LEFT = 'i';
    const LABEL_OUTSIDEEND = 'outEnd';
    const LABEL_RIGHT = 'r';
    const LABEL_TOP = 't';

    /**
     * Title
     *
     * @var string
     */
    private $title = 'Series Title';

    /**
     * Fill
     *
     * @var PHPPowerPoint_Style_Fill
     */
    private $fill;

    /**
     * Values (key/value)
     *
     * @var array
     */
    private $values = array();

    /**
     * DataPointFills (key/value)
     *
     * @var array
     */
    private $dataPointFills = array();

    /**
     * ShowSeriesName
     *
     * @var boolean
     */
    private $showSeriesName = false;

    /**
     * ShowCategoryName
     *
     * @var boolean
     */
    private $showCategoryName = false;

    /**
     * ShowValue
     *
     * @var boolean
     */
    private $showValue = true;

    /**
     * ShowPercentage
     *
     * @var boolean
     */
    private $showPercentage = false;

    /**
     * ShowLeaderLines
     *
     * @var boolean
     */
    private $showLeaderLines = true;

    /**
     * Font
     *
     * @var PHPPowerPoint_Style_Font
     */
    private $font;

    /**
     * Label position
     *
     * @var string
     */
    private $labelPosition = 'ctr';

    /**
     * Create a new PHPPowerPoint_Shape_Chart_Series instance
     *
     * @param string $title  Title
     * @param array  $values Values
     */
    public function __construct($title = 'Series Title', $values = array())
    {
        $this->fill = new Fill();
        $this->font = new Font();
        $this->font->setName('Calibri');
        $this->font->setSize(9);
        $this->title  = $title;
        $this->values = $values;
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
     * @param  string                           $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function setTitle($value = 'Series Title')
    {
        $this->title = $value;

        return $this;
    }

    /**
     * Get Fill
     *
     * @return PHPPowerPoint_Style_Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Get DataPointFill
     *
     * @param  int                      $dataPointIndex Data point index.
     * @return PHPPowerPoint_Style_Fill
     */
    public function getDataPointFill($dataPointIndex)
    {
        if (!isset($this->dataPointFills[$dataPointIndex])) {
            $this->dataPointFills[$dataPointIndex] = new Fill();
        }

        return $this->dataPointFills[$dataPointIndex];
    }

    /**
     * Get DataPointFills
     *
     * @return array
     */
    public function getDataPointFills()
    {
        return $this->dataPointFills;
    }

    /**
     * Get Values
     *
     * @return array
     */
    public function getValues()
    {
        return $this->values;
    }

    /**
     * Set Values
     *
     * @param  array                            $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function setValues($value = array())
    {
        $this->values = $value;

        return $this;
    }

    /**
     * Add Value
     *
     * @param  mixed                            $key
     * @param  mixed                            $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function addValue($key, $value)
    {
        $this->values[$key] = $value;

        return $this;
    }

    /**
     * Get ShowSeriesName
     *
     * @return boolean
     */
    public function hasShowSeriesName()
    {
        return $this->showSeriesName;
    }

    /**
     * Set ShowSeriesName
     *
     * @param  boolean                          $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function setShowSeriesName($value)
    {
        $this->showSeriesName = $value;

        return $this;
    }

    /**
     * Get ShowCategoryName
     *
     * @return boolean
     */
    public function hasShowCategoryName()
    {
        return $this->showCategoryName;
    }

    /**
     * Set ShowCategoryName
     *
     * @param  boolean                          $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function setShowCategoryName($value)
    {
        $this->showCategoryName = $value;

        return $this;
    }

    /**
     * Get ShowValue
     *
     * @return boolean
     */
    public function hasShowValue()
    {
        return $this->showValue;
    }

    /**
     * Set ShowValue
     *
     * @param  boolean                          $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function setShowValue($value)
    {
        $this->showValue = $value;

        return $this;
    }

    /**
     * Get ShowPercentage
     *
     * @return boolean
     */
    public function hasShowPercentage()
    {
        return $this->showPercentage;
    }

    /**
     * Set ShowPercentage
     *
     * @param  boolean                          $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function setShowPercentage($value)
    {
        $this->showPercentage = $value;

        return $this;
    }

    /**
     * Get ShowLeaderLines
     *
     * @return boolean
     */
    public function hasShowLeaderLines()
    {
        return $this->showLeaderLines;
    }

    /**
     * Set ShowLeaderLines
     *
     * @param  boolean                          $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function setShowLeaderLines($value)
    {
        $this->showLeaderLines = $value;

        return $this;
    }

    /**
     * Get font
     *
     * @return PHPPowerPoint_Style_Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param  PHPPowerPoint_Style_Font               $pFont Font
     * @throws \Exception
     * @return PHPPowerPoint_Shape_RichText_Paragraph
     */
    public function setFont(Font $pFont = null)
    {
        $this->font = $pFont;

        return $this;
    }

    /**
     * Get label position
     *
     * @return string
     */
    public function getLabelPosition()
    {
        return $this->labelPosition;
    }

    /**
     * Set label position
     *
     * @param  string                           $value
     * @return PHPPowerPoint_Shape_Chart_Series
     */
    public function setLabelPosition($value)
    {
        $this->labelPosition = $value;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5((is_null($this->fill) ? 'null' : $this->fill->getHashCode()) . (is_null($this->font) ? 'null' : $this->font->getHashCode()) . var_export($this->values, true) . var_export($this, true) . __CLASS__);
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
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
}
