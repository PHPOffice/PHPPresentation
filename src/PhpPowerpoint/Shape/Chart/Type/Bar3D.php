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

namespace PhpOffice\PhpPowerpoint\Shape\Chart\Type;

use PhpOffice\PhpPowerpoint\IComparable;
use PhpOffice\PhpPowerpoint\Shape\Chart\Type;
use PhpOffice\PhpPowerpoint\Shape\Chart\Series;

/**
 * PHPPowerPoint_Shape_Chart_Type_Bar3D
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape_Chart_Type
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Bar3D extends Type implements IComparable
{
    /**
     * Data
     *
     * @var array
     */
    private $data = array();

    /**
     * Get Data
     *
     * @return array
     */
    public function getData()
    {
        return $this->data;
    }

    /**
     * Set Data
     *
     * @param  array                          $value Array of PHPPowerPoint_Shape_Chart_Series
     * @return PHPPowerPoint_Shape_Type_Bar3D
     */
    public function setData($value = array())
    {
        $this->data = $value;

        return $this;
    }

    /**
     * Add Series
     *
     * @param  PHPPowerPoint_Shape_Chart_Series $value
     * @return PHPPowerPoint_Shape_Type_Bar3D
     */
    public function addSeries(Series $value)
    {
        $this->data[] = $value;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        $hash = '';
        foreach ($this->data as $series) {
            $hash .= $series->getHashCode();
        }

        return md5($hash . __CLASS__);
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
