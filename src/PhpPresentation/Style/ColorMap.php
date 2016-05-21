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
 * PhpOffice\PhpPresentation\Style\ColorMap
 */
class ColorMap
{
    /**
     * Mapping - Stores the mapping betweenSlide and theme
     *
     * @var array
     */
    protected $mapping = array();

    /**
     * ColorMap constructor.
     * Create a new ColorMap with standard values
     */
    public function __construct()
    {
        $this->mapping = array("bg1" => "lt1",
            "tx1" => "dk1",
            "bg2" => "lt2",
            "tx2" => "dk2",
            "accent1" => "accent1",
            "accent2" => "accent2",
            "accent3" => "accent3",
            "accent4" => "accent4",
            "accent5" => "accent5",
            "accent6" => "accent6",
            "hlink" => "hlink",
            "folHlink" => "folHlink");
    }

    /**
     * Change the color of one of the elements in the map
     *
     * @param string $item
     * @param string $newThemeColor
     */
    public function changeColor($item, $newThemeColor)
    {
        $this->mapping[$item] = $newThemeColor;
    }

    /**
     * Store a new map. For use with the reader
     *
     * @param $newMappingArray
     */
    public function setNewMapping($newMappingArray)
    {
        $this->mapping = $newMappingArray;
    }

    /**
     * Get the whole mapping as an array
     *
     * @return array
     */
    public function getMapping()
    {
        return $this->mapping;
    }
}