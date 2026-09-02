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
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;

class Series implements ComparableInterface
{
    // Label positions
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
     * DataPointFills (key/value).
     *
     * @var array<int, Fill>
     */
    protected $dataPointFills = [];

    /**
     * DataPointOutlines (key/value).
     *
     * @var array<int, Outline>
     */
    protected $dataPointOutlines = [];

    /**
     * Data Label Number Format.
     *
     * @var string
     */
    protected $dlblNumFormat = '';

    /**
     * @var null|string
     */
    protected $separator;

    /**
     * @var Fill
     */
    protected $fill;

    /**
     * @var Font
     */
    protected $font;

    /**
     * The fill behind the data labels, not behind the series itself.
     *
     * @var Fill
     */
    protected $labelFill;

    /**
     * @var string
     */
    protected $labelPosition = 'ctr';

    /**
     * @var Marker
     */
    protected $marker;

    /**
     * @var null|Outline
     */
    protected $outline;

    /**
     * Show Category Name.
     *
     * @var bool
     */
    private $showCategoryName = false;

    /**
     * Show Leader Lines.
     *
     * @var bool
     */
    private $showLeaderLines = true;

    /**
     * Show Legend Key.
     *
     * @var bool
     */
    private $showLegendKey = false;

    /**
     * ShowPercentage.
     *
     * @var bool
     */
    private $showPercentage = false;

    /**
     * ShowSeriesName.
     *
     * @var bool
     */
    private $showSeriesName = false;

    /**
     * ShowValue.
     *
     * @var bool
     */
    private $showValue = true;

    /**
     * Title.
     *
     * @var string
     */
    private $title = 'Series Title';

    /**
     * Values (key/value).
     *
     * @var array<array-key, null|string>
     */
    private $values = [];

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * @param array<array-key, null|string> $values
     */
    public function __construct(string $title = 'Series Title', array $values = [])
    {
        $this->fill = new Fill();
        // the labels are born with no fill of their own: the application draws them as it always did
        $this->labelFill = (new Fill())->setFillType(Fill::FILL_UNSET);
        $this->font = new Font();
        $this->font->setName('Calibri');
        $this->font->setSize(9);
        $this->marker = new Marker();

        $this->title = $title;
        $this->values = $values;
    }

    /**
     * Get Title.
     */
    public function getTitle(): string
    {
        return $this->title;
    }

    /**
     * Set Title.
     */
    public function setTitle(string $value = 'Series Title'): self
    {
        $this->title = $value;

        return $this;
    }

    /**
     * Get Data Label NumFormat.
     */
    public function getDlblNumFormat(): string
    {
        return $this->dlblNumFormat;
    }

    /**
     * Has Data Label NumFormat.
     */
    public function hasDlblNumFormat(): bool
    {
        return !empty($this->dlblNumFormat);
    }

    /**
     * Set Data Label NumFormat.
     */
    public function setDlblNumFormat(string $value = ''): self
    {
        $this->dlblNumFormat = $value;

        return $this;
    }

    public function getFill(): Fill
    {
        return $this->fill;
    }

    /**
     * A series is born with a fill, and asking for it to be taken away replaces it with one of type
     * `Fill::FILL_UNSET` rather than with a null: every writer reads it without asking whether it is
     * there. A series that never named a fill leaves the colour to the application, which is what a
     * null used to mean.
     *
     * @param null|Fill $fill passing `null` is deprecated, pass a `Fill` of type `Fill::FILL_UNSET`
     */
    public function setFill(?Fill $fill = null): self
    {
        $this->fill = $fill ?? (new Fill())->setFillType(Fill::FILL_UNSET);

        return $this;
    }

    /**
     * The fill drawn behind the data labels. A series that never named one carries a fill of type
     * `Fill::FILL_UNSET`, which is written as nothing at all.
     */
    public function getLabelFill(): Fill
    {
        return $this->labelFill;
    }

    /**
     * A semi-transparent plate under the labels is a `Fill::FILL_SOLID` whose colour carries an
     * alpha, e.g. `new Color('D6FFFFFF')`. Asking for the fill to be taken away replaces it with
     * one of type `Fill::FILL_UNSET` rather than with a null, as `setFill()` does.
     */
    public function setLabelFill(?Fill $fill = null): self
    {
        $this->labelFill = $fill ?? (new Fill())->setFillType(Fill::FILL_UNSET);

        return $this;
    }

    /**
     * @param int $dataPointIndex data point index
     */
    public function getDataPointFill(int $dataPointIndex): Fill
    {
        if (!isset($this->dataPointFills[$dataPointIndex])) {
            $this->dataPointFills[$dataPointIndex] = new Fill();
        }

        return $this->dataPointFills[$dataPointIndex];
    }

    /**
     * @param int $dataPointIndex data point index
     */
    public function setDataPointFill(int $dataPointIndex, Fill $fill): self
    {
        $this->dataPointFills[$dataPointIndex] = $fill;

        return $this;
    }

    /**
     * @return Fill[]
     */
    public function getDataPointFills(): array
    {
        return $this->dataPointFills;
    }

    /**
     * @param int $dataPointIndex data point index
     */
    public function getDataPointOutline(int $dataPointIndex): Outline
    {
        if (!isset($this->dataPointOutlines[$dataPointIndex])) {
            $this->dataPointOutlines[$dataPointIndex] = new Outline();
        }

        return $this->dataPointOutlines[$dataPointIndex];
    }

    /**
     * @param int $dataPointIndex data point index
     */
    public function setDataPointOutline(int $dataPointIndex, Outline $outline): self
    {
        $this->dataPointOutlines[$dataPointIndex] = $outline;

        return $this;
    }

    /**
     * @return Outline[]
     */
    public function getDataPointOutlines(): array
    {
        return $this->dataPointOutlines;
    }

    /**
     * The indexes of the data points carrying a fill or an outline of their own.
     *
     * @return array<int, int>
     */
    public function getDataPointIndexes(): array
    {
        $indexes = array_keys($this->dataPointFills + $this->dataPointOutlines);
        sort($indexes);

        return $indexes;
    }

    /**
     * Get Values.
     *
     * @return array<array-key, null|string>
     */
    public function getValues(): array
    {
        return $this->values;
    }

    /**
     * Set Values.
     *
     * @param array<array-key, null|string> $values
     */
    public function setValues(array $values = []): self
    {
        $this->values = $values;

        return $this;
    }

    /**
     * Add Value.
     */
    public function addValue(string $key, ?string $value): self
    {
        $this->values[$key] = $value;

        return $this;
    }

    /**
     * Get ShowSeriesName.
     */
    public function hasShowSeriesName(): bool
    {
        return $this->showSeriesName;
    }

    /**
     * Set ShowSeriesName.
     */
    public function setShowSeriesName(bool $value): self
    {
        $this->showSeriesName = $value;

        return $this;
    }

    /**
     * Get ShowCategoryName.
     */
    public function hasShowCategoryName(): bool
    {
        return $this->showCategoryName;
    }

    /**
     * Set ShowCategoryName.
     */
    public function setShowCategoryName(bool $value): self
    {
        $this->showCategoryName = $value;

        return $this;
    }

    /**
     * Get ShowValue.
     */
    public function hasShowLegendKey(): bool
    {
        return $this->showLegendKey;
    }

    /**
     * Set ShowValue.
     */
    public function setShowLegendKey(bool $value): self
    {
        $this->showLegendKey = $value;

        return $this;
    }

    /**
     * Get ShowValue.
     */
    public function hasShowValue(): bool
    {
        return $this->showValue;
    }

    /**
     * Set ShowValue.
     */
    public function setShowValue(bool $value): self
    {
        $this->showValue = $value;

        return $this;
    }

    /**
     * Get ShowPercentage.
     */
    public function hasShowPercentage(): bool
    {
        return $this->showPercentage;
    }

    /**
     * Set ShowPercentage.
     */
    public function setShowPercentage(bool $value): self
    {
        $this->showPercentage = $value;

        return $this;
    }

    public function hasShowSeparator(): bool
    {
        return null !== $this->separator;
    }

    public function setSeparator(?string $pValue): self
    {
        $this->separator = $pValue;

        return $this;
    }

    public function getSeparator(): ?string
    {
        return $this->separator;
    }

    /**
     * Get ShowLeaderLines.
     */
    public function hasShowLeaderLines(): bool
    {
        return $this->showLeaderLines;
    }

    /**
     * Set ShowLeaderLines.
     *
     * @param bool $value
     *
     * @return self
     */
    public function setShowLeaderLines($value)
    {
        $this->showLeaderLines = $value;

        return $this;
    }

    /**
     * Get font.
     */
    public function getFont(): Font
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
        $this->font = $pFont ?? new Font();

        return $this;
    }

    /**
     * Get label position.
     */
    public function getLabelPosition(): string
    {
        return $this->labelPosition;
    }

    /**
     * Set label position.
     */
    public function setLabelPosition(string $value): self
    {
        $this->labelPosition = $value;

        return $this;
    }

    public function getMarker(): Marker
    {
        return $this->marker;
    }

    public function setMarker(Marker $marker): self
    {
        $this->marker = $marker;

        return $this;
    }

    public function getOutline(): ?Outline
    {
        return $this->outline;
    }

    public function setOutline(?Outline $outline): self
    {
        $this->outline = $outline;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5((null === $this->fill ? 'null' : $this->fill->getHashCode()) . (null === $this->font ? 'null' : $this->font->getHashCode()) . var_export($this->values, true) . var_export($this, true) . __CLASS__);
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
     */
    public function setHashIndex(int $value): self
    {
        $this->hashIndex = $value;

        return $this;
    }

    /**
     * @see http://php.net/manual/en/language.oop5.cloning.php
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
