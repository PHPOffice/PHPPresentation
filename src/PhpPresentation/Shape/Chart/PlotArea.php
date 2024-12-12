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
use PhpOffice\PhpPresentation\Exception\UndefinedChartTypeException;
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\PlotArea.
 */
class PlotArea implements ComparableInterface
{
    /**
     * Type.
     *
     * @var null|AbstractType
     */
    private $type;

    /**
     * Axis X.
     *
     * @var Axis
     */
    private $axisX;

    /**
     * Axis Y.
     *
     * @var Axis
     */
    private $axisY;

    /**
     * OffsetX (as a fraction of the chart).
     *
     * @var float
     */
    private $offsetX = 0;

    /**
     * OffsetY (as a fraction of the chart).
     *
     * @var float
     */
    private $offsetY = 0;

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

    public function __construct()
    {
        $this->axisX = new Axis();
        $this->axisY = new Axis();
    }

    public function __clone()
    {
        $this->axisX = clone $this->axisX;
        $this->axisY = clone $this->axisY;
    }

    public function getType(): AbstractType
    {
        if (null === $this->type) {
            throw new UndefinedChartTypeException();
        }

        return $this->type;
    }

    public function setType(AbstractType $value): self
    {
        $this->type = $value;

        return $this;
    }

    /**
     * Get Axis X.
     */
    public function getAxisX(): Axis
    {
        return $this->axisX;
    }

    /**
     * Get Axis Y.
     */
    public function getAxisY(): Axis
    {
        return $this->axisY;
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
    public function setOffsetX(float $pValue = 0): self
    {
        $this->offsetX = $pValue;

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
    public function setOffsetY(float $pValue = 0): self
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
    public function setWidth(int $pValue = 0): self
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
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5((null === $this->type ? 'null' : $this->type->getHashCode()) . $this->axisX->getHashCode() . $this->axisY->getHashCode() . $this->offsetX . $this->offsetY . $this->width . $this->height . __CLASS__);
    }

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

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
     * @return PlotArea
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;

        return $this;
    }
}
