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
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Series
 */
class Series implements ComparableInterface
{
    /* Label positions */
    const LABEL_BESTFIT = 'bestFit';
    const LABEL_BOTTOM = 'b';
    const LABEL_CENTER = 'ctr';
    const LABEL_INSIDEBASE = 'inBase';
    const LABEL_INSIDEEND = 'inEnd';
    const LABEL_LEFT = 'i';
    const LABEL_OUTSIDEEND = 'outEnd';
    const LABEL_RIGHT = 'r';
    const LABEL_TOP = 't';

    /**
     * DataPointFills (key/value)
     *
     * @var array
     */
    private $dataPointFills = array();

    /**
     * Data Label Number Format
     *
     * @var string
     */
    private $DlblNumFormat = '';

    /**
     * Fill
     *
     * @var \PhpOffice\PhpPresentation\Style\Fill
     */
    private $fill;

    /**
     * Font
     *
     * @var \PhpOffice\PhpPresentation\Style\Font
     */
    private $font;

    /**
     * Label position
     *
     * @var string
     */
    private $labelPosition = 'ctr';

    /**
     * @var Marker
     */
    protected $marker;

    /**
     * @var Outline
     */
    protected $outline;

    /**
     * ShowCategoryName
     *
     * @var boolean
     */
    private $showCategoryName = false;

    /**
     * ShowLeaderLines
     *
     * @var boolean
     */
    private $showLeaderLines = true;

    /**
     * ShowPercentage
     *
     * @var boolean
     */
    private $showPercentage = false;

    /**
     * ShowSeriesName
     *
     * @var boolean
     */
    private $showSeriesName = false;

    /**
     * ShowValue
     *
     * @var boolean
     */
    private $showValue = true;

    /**
     * Title
     *
     * @var string
     */
    private $title = 'Series Title';

    /**
     * Values (key/value)
     *
     * @var array
     */
    private $values = array();

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Chart\Series instance
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
        $this->marker = new Marker();
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
     */
    public function setTitle($value = 'Series Title')
    {
        $this->title = $value;

        return $this;
    }

    /**
     * Get Data Label NumFormat
     *
     * @return string
     */
    public function getDlblNumFormat()
    {
        return $this->DlblNumFormat;
    }

    /**
     * Has Data Label NumFormat
     *
     * @return string
     */
    public function hasDlblNumFormat()
    {
        return !empty($this->DlblNumFormat);
    }

    /**
     * Set Data Label NumFormat
     *
     * @param  string $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
     */
    public function setDlblNumFormat($value = '')
    {
        $this->DlblNumFormat = $value;
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
     * @return Series
     */
    public function setFill(Fill $fill = null)
    {
        $this->fill = $fill;
        return $this;
    }

    /**
     * Get DataPointFill
     *
     * @param  int                      $dataPointIndex Data point index.
     * @return \PhpOffice\PhpPresentation\Style\Fill
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
     * @return Fill[]
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
     */
    public function setShowLeaderLines($value)
    {
        $this->showLeaderLines = $value;

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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
     */
    public function setLabelPosition($value)
    {
        $this->labelPosition = $value;

        return $this;
    }

    /**
     * @return Marker
     */
    public function getMarker()
    {
        return $this->marker;
    }

    /**
     * @param Marker $marker
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
     */
    public function setMarker(Marker $marker)
    {
        $this->marker = $marker;
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
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
        return md5((is_null($this->fill) ? 'null' : $this->fill->getHashCode()) . (is_null($this->font) ? 'null' : $this->font->getHashCode()) . var_export($this->values, true) . var_export($this, true) . __CLASS__);
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Series
     */
    public function setHashIndex($value)
    {
        $this->hashIndex = $value;
        return $this;
    }

    /**
     * @return mixed
     * @link http://php.net/manual/en/language.oop5.cloning.php
     */
    public function __clone()
    {
        $this->font = clone $this->font;
        $this->marker = clone $this->marker;
        if (is_object($this->outline)) {
            $this->outline = clone $this->outline;
        }
    }
}
