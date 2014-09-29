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

/**
 * \PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D
 */
class Bar3D extends AbstractType implements ComparableInterface
{
    /** Orientation of bars */
    const DIRECTION_VERTICAL = 'col';
    const DIRECTION_HORIZONTAL = 'bar';

    /**
     * Orientation of bars
     *
     * @var string
     */
    protected $barDirection = self::DIRECTION_VERTICAL;

    /**
     * Set bar orientation
     *
     * @param string                          $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D
     */
    public function setBarDirection($value = self::DIRECTION_VERTICAL)
    {
        $this->barDirection = $value;
        return $this;
    }

    /**
     * Get orientation
     *
     * @return string
     */
    public function getBarDirection()
    {
        return $this->barDirection;
    }
    
    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        $hash = '';
        foreach ($this->getData() as $series) {
            $hash .= $series->getHashCode();
        }
        return md5($hash . __CLASS__);
    }
}
