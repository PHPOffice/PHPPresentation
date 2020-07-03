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
use PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\PlotArea
 */
class PlotArea implements ComparableInterface
{
    /**
     * Type
     *
     * @var \PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType
     */
    private $type;

    /**
     * Axis X
     *
     * @var \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    private $axisX;

    /**
     * Axis Y
     *
     * @var \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    private $axisY;

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
     * Create a new \PhpOffice\PhpPresentation\Shape\Chart\PlotArea instance
     */
    public function __construct()
    {
        $this->type  = null;
        $this->axisX = new Axis();
        $this->axisY = new Axis();
    }
    
    public function __clone()
    {
        $this->axisX     = clone $this->axisX;
        $this->axisY     = clone $this->axisY;
    }

    /**
     * Get type
     *
     * @return AbstractType
     * @throws \Exception
     */
    public function getType()
    {
        if (is_null($this->type)) {
            throw new \Exception('Chart type has not been set.');
        }

        return $this->type;
    }

    /**
     * Set type
     *
     * @param \PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractType $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\PlotArea
     */
    public function setType(Type\AbstractType $value)
    {
        $this->type = $value;

        return $this;
    }

    /**
     * Get Axis X
     *
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    public function getAxisX()
    {
        return $this->axisX;
    }

    /**
     * Get Axis Y
     *
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    public function getAxisY()
    {
        return $this->axisY;
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
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
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Title
     */
    public function setHeight($value = 0)
    {
        $this->height = (double)$value;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5((is_null($this->type) ? 'null' : $this->type->getHashCode()) . $this->axisX->getHashCode() . $this->axisY->getHashCode() . $this->offsetX . $this->offsetY . $this->width . $this->height . __CLASS__);
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
