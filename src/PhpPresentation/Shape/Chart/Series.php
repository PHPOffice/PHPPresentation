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

class Series implements ComparableInterface
{
    /* Label positions */
    public const LABEL_BESTFIT = 'bestFit';
    public const LABEL_BOTTOM = 'b';
    public const LABEL_CENTER = 'ctr';
    public const LABEL_INSIDEBASE = 'inBase';
    public const LABEL_INSIDEEND = 'inEnd';
    public const LABEL_LEFT = 'i';
    public const LABEL_OUTSIDEEND = 'outEnd';
    public const LABEL_RIGHT = 'r';
    public const LABEL_TOP = 't';

    /**
     * DataPointFills (key/value)
     * @var array<int, Fill>
     */
    protected $dataPointFills = array();

    /**
     * Data Label Number Format
     * @var string
     */
    protected $DlblNumFormat = '';

    /**
     * @var string|null
     */
    protected $separator;

    /**
     * @var Fill|null
     */
    protected $fill;

    /**
     * @var Font|null
     */
    protected $font;

    /**
     * @var string
     */
    protected $labelPosition = 'ctr';

    /**
     * @var Marker
     */
    protected $marker;

    /**
     * @var Outline|null
     */
    protected $outline;

    /**
     * Show Category Name
     * @var boolean
     */
    private $showCategoryName = false;

    /**
     * Show Leader Lines
     * @var boolean
     */
    private $showLeaderLines = true;

    /**
     * Show Legend Key
     * @var bool
     */
    private $showLegendKey = false;

    /**
     * ShowPercentage
     * @var boolean
     */
    private $showPercentage = false;

    /**
     * ShowSeriesName
     * @var boolean
     */
    private $showSeriesName = false;

    /**
     * ShowValue
     * @var boolean
     */
    private $showValue = true;

    /**
     * Title
     * @var string
     */
    private $title = 'Series Title';

    /**
     * Values (key/value)
     * @var array<string, string>
     */
    private $values = array();

    /**
     * Hash index
     * @var int
     */
    private $hashIndex;

    /**
     * @param string $title
     * @param array<string, string> $values
     */
    public function __construct(string $title = 'Series Title', array $values = array())
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
    public function getTitle(): string
    {
        return $this->title;
    }

    /**
     * Set Title
     *
     * @param string $value
     * @return self
     */
    public function setTitle(string $value = 'Series Title'): self
    {
        $this->title = $value;

        return $this;
    }

    /**
     * Get Data Label NumFormat
     *
     * @return string
     */
    public function getDlblNumFormat(): string
    {
        return $this->DlblNumFormat;
    }

    /**
     * Has Data Label NumFormat
     *
     * @return bool
     */
    public function hasDlblNumFormat(): bool
    {
        return !empty($this->DlblNumFormat);
    }

    /**
     * Set Data Label NumFormat
     *
     * @param string $value
     * @return self
     */
    public function setDlblNumFormat(string $value = ''): self
    {
        $this->DlblNumFormat = $value;
        return $this;
    }

    /**
     * @return Fill
     */
    public function getFill(): ?Fill
    {
        return $this->fill;
    }

    /**
     * @param Fill|null $fill
     * @return self
     */
    public function setFill(Fill $fill = null): self
    {
        $this->fill = $fill;
        return $this;
    }

    /**
     * @param int $dataPointIndex Data point index.
     * @return Fill
     */
    public function getDataPointFill(int $dataPointIndex): Fill
    {
        if (!isset($this->dataPointFills[$dataPointIndex])) {
            $this->dataPointFills[$dataPointIndex] = new Fill();
        }

        return $this->dataPointFills[$dataPointIndex];
    }

    /**
     * @return Fill[]
     */
    public function getDataPointFills(): array
    {
        return $this->dataPointFills;
    }

    /**
     * Get Values
     *
     * @return array<string, string>
     */
    public function getValues(): array
    {
        return $this->values;
    }

    /**
     * Set Values
     *
     * @param array<string, string> $values
     * @return self
     */
    public function setValues(array $values = array()): self
    {
        $this->values = $values;

        return $this;
    }

    /**
     * Add Value
     *
     * @param string $key
     * @param string $value
     * @return self
     */
    public function addValue(string $key, string $value): self
    {
        $this->values[$key] = $value;

        return $this;
    }

    /**
     * Get ShowSeriesName
     *
     * @return boolean
     */
    public function hasShowSeriesName(): bool
    {
        return $this->showSeriesName;
    }

    /**
     * Set ShowSeriesName
     *
     * @param boolean $value
     * @return self
     */
    public function setShowSeriesName(bool $value): self
    {
        $this->showSeriesName = $value;

        return $this;
    }

    /**
     * Get ShowCategoryName
     *
     * @return boolean
     */
    public function hasShowCategoryName(): bool
    {
        return $this->showCategoryName;
    }

    /**
     * Set ShowCategoryName
     *
     * @param boolean $value
     * @return self
     */
    public function setShowCategoryName(bool $value): self
    {
        $this->showCategoryName = $value;

        return $this;
    }

    /**
     * Get ShowValue
     *
     * @return bool
     */
    public function hasShowLegendKey(): bool
    {
        return $this->showLegendKey;
    }

    /**
     * Set ShowValue
     *
     * @param bool $value
     * @return self
     */
    public function setShowLegendKey(bool $value): self
    {
        $this->showLegendKey = $value;

        return $this;
    }

    /**
     * Get ShowValue
     *
     * @return boolean
     */
    public function hasShowValue(): bool
    {
        return $this->showValue;
    }

    /**
     * Set ShowValue
     *
     * @param boolean $value
     * @return self
     */
    public function setShowValue(bool $value): self
    {
        $this->showValue = $value;

        return $this;
    }

    /**
     * Get ShowPercentage
     *
     * @return boolean
     */
    public function hasShowPercentage(): bool
    {
        return $this->showPercentage;
    }

    /**
     * Set ShowPercentage
     *
     * @param boolean $value
     * @return self
     */
    public function setShowPercentage(bool $value): self
    {
        $this->showPercentage = $value;

        return $this;
    }

    /**
     * @return boolean
     */
    public function hasShowSeparator(): bool
    {
        return !is_null($this->separator);
    }

    /**
     * @param string|null $pValue
     * @return self
     */
    public function setSeparator(?string $pValue): self
    {
        $this->separator = $pValue;
        return $this;
    }

    /**
     * @return string|null
     */
    public function getSeparator(): ?string
    {
        return $this->separator;
    }

    /**
     * Get ShowLeaderLines
     *
     * @return boolean
     */
    public function hasShowLeaderLines(): bool
    {
        return $this->showLeaderLines;
    }

    /**
     * Set ShowLeaderLines
     *
     * @param boolean                          $value
     * @return self
     */
    public function setShowLeaderLines($value)
    {
        $this->showLeaderLines = $value;

        return $this;
    }

    /**
     * Get font
     *
     * @return Font
     */
    public function getFont(): ?Font
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param Font|null $pFont Font
     * @return self
     */
    public function setFont(Font $pFont = null): self
    {
        $this->font = $pFont;

        return $this;
    }

    /**
     * Get label position
     *
     * @return string
     */
    public function getLabelPosition(): string
    {
        return $this->labelPosition;
    }

    /**
     * Set label position
     *
     * @param string $value
     * @return self
     */
    public function setLabelPosition(string $value): self
    {
        $this->labelPosition = $value;

        return $this;
    }

    /**
     * @return Marker
     */
    public function getMarker(): Marker
    {
        return $this->marker;
    }

    /**
     * @param Marker $marker
     * @return self
     */
    public function setMarker(Marker $marker): self
    {
        $this->marker = $marker;
        return $this;
    }

    /**
     * @return Outline|null
     */
    public function getOutline(): ?Outline
    {
        return $this->outline;
    }

    /**
     * @param Outline|null $outline
     * @return self
     */
    public function setOutline(?Outline $outline): self
    {
        $this->outline = $outline;
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5((is_null($this->fill) ? 'null' : $this->fill->getHashCode()) . (is_null($this->font) ? 'null' : $this->font->getHashCode()) . var_export($this->values, true) . var_export($this, true) . __CLASS__);
    }

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
     * @return self
     */
    public function setHashIndex(int $value): self
    {
        $this->hashIndex = $value;
        return $this;
    }


    /**
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
