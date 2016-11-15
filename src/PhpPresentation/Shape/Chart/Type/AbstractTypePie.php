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

namespace PhpOffice\PhpPresentation\Shape\Chart\Type;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Type\Bar
 */
class AbstractTypePie extends AbstractType
{
    /**
     * Create a new self instance
     */
    public function __construct()
    {
        $this->hasAxisX = false;
        $this->hasAxisY = false;
    }
    
    /**
     * Explosion of the Pie
     *
     * @var integer
     */
    protected $explosion = 0;
    
    /**
     * Set explosion
     *
     * @param integer $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Type\AbstractTypePie
     */
    public function setExplosion($value = 0)
    {
        $this->explosion = $value;
        return $this;
    }
    
    /**
     * Get orientation
     *
     * @return string
     */
    public function getExplosion()
    {
        return $this->explosion;
    }
    
    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        $hash = '';
        foreach ($this->getSeries() as $series) {
            $hash .= $series->getHashCode();
        }
        return $hash;
    }
}
