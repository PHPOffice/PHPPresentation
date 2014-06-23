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

use PhpOffice\PhpPowerpoint\ComparableInterface;
use PhpOffice\PhpPowerpoint\Shape\Chart\AbstractType;
use PhpOffice\PhpPowerpoint\Shape\Chart\Series;

/**
 * \PhpOffice\PhpPowerpoint\Shape\Chart\Type\Scatter
 */
class Scatter extends AbstractType implements ComparableInterface
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
     * @param  array $value Array of \PhpOffice\PhpPowerpoint\Shape\Chart\Series
     * @return self
     */
    public function setData($value = array())
    {
        $this->data = $value;

        return $this;
    }

    /**
     * Add Series
     *
     * @param  \PhpOffice\PhpPowerpoint\Shape\Chart\Series $value
     * @return self
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
}
