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
use PhpOffice\PhpPowerpoint\Shape\Chart\Type\AbstractType;

/**
 * \PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea
 */
class PlotArea implements ComparableInterface
{
    /**
     * Type
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\Chart\AbstractType
     */
    private $type;

    /**
     * Axis X
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\Chart\Axis
     */
    private $axisX;

    /**
     * Axis Y
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\Chart\Axis
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
     * Create a new \PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea instance
     */
    public function __construct()
    {
        $this->type  = null;
        $this->axisX = new Axis();
        $this->axisY = new Axis();
    }

    /**
     * Get type
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\AbstractType
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
     * @param \PhpOffice\PhpPowerpoint\Shape\Chart\AbstractType $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea
     */
    public function setType(AbstractType $value)
    {
        $this->type = $value;

        return $this;
    }

    /**
     * Get Axis X
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Axis
     */
    public function getAxisX()
    {
        return $this->axisX;
    }

    /**
     * Get Axis Y
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Axis
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Title
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Title
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Title
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Title
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
