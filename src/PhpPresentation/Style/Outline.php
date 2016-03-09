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

namespace PhpOffice\PhpPresentation\Style;

/**
 * \PhpOffice\PhpPresentation\Style\Outline
 */
class Outline
{
    /**
     * @var Fill
     */
    protected $fill;
    /**
     * @var int
     */
    protected $width;

    /**
     * @return void
     */
    public function __construct()
    {
        $this->fill = new Fill();
    }

    /**
     * @return Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * @param Fill $fill
     * @return Outline
     */
    public function setFill(Fill $fill)
    {
        $this->fill = $fill;
        return $this;
    }

    /**
     * @return int
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Value in points
     * @param int $width
     * @return Outline
     */
    public function setWidth($width)
    {
        $this->width = intval($width);
        return $this;
    }
}
