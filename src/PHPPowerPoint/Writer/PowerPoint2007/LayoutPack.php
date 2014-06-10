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

namespace PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;

/**
 * PHPPowerPoint_Writer_PowerPoint2007_LayoutPack
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
abstract class LayoutPack
{
    /**
     * Master slides
     *
     * Structure:
     * - masterid
     * - body
     *
     * @var array
     */
    protected $_masterSlides = array();

    /**
     * Master slide relations
     *
     * Structure:
     * - master id
     * - id (relation id)
     * - type
     * - contentType
     * - target (full path in OpenXML package)
     * - contents (body)
     *
     * @var array
     */
    protected $_masterSlideRelations = array();

    /**
     * Themes
     *
     * Structure:
     * - masterid
     * - body
     *
     * @var array
     */
    protected $_themes = '';

    /**
     * Theme relations
     *
     * Structure:
     * - masterid
     * - id (relation id)
     * - type
     * - contentType
     * - target (full path in OpenXML package)
     * - contents (body)
     *
     * @var array
     */
    protected $_themeRelations = array();

    /**
     * Array of slide layouts.
     *
     * These are all an array consisting of:
     * - masterid (int)
     * - name (string)
     * - body (string)
     *
     * @var array
     */
    protected $_layouts = array();

    /**
     * Layout relations
     *
     * Structure:
     * - layoutId (referencing layout id in layouts array)
     * - id (relation id)
     * - type
     * - contentType
     * - target (full path in OpenXML package)
     * - contents (body)
     *
     * @var array
     */
    protected $_layoutRelations = array();

    /**
     * Get master slides
     *
     * @return array
     */
    public function getMasterSlides()
    {
        return $this->_masterSlides;
    }

    /**
     * Get master slide relations
     *
     * @return array
     */
    public function getMasterSlideRelations()
    {
        return $this->_masterSlideRelations;
    }

    /**
     * Get themes
     *
     * @return array
     */
    public function getThemes()
    {
        return $this->_themes;
    }

    /**
     * Get theme relations
     *
     * @return array
     */
    public function getThemeRelations()
    {
        return $this->_themeRelations;
    }

    /**
     * Get array of slide layouts
     *
     * @return array
     */
    public function getLayouts()
    {
        return $this->_layouts;
    }

    /**
     * Get array of slide layout relations
     *
     * @return array
     */
    public function getLayoutRelations()
    {
        return $this->_layoutRelations;
    }

    /**
     * Find specific slide layout.
     *
     * This is an array consisting of:
     * - masterid
     * - name (string)
     * - body (string)
     *
     * @param string $name
     * @param int $masterId
     * @return array
     * @throws Exception
     */
    public function findLayout($name = '', $masterId = 1)
    {
        foreach ($this->_layouts as $layout) {
            if ($layout['name'] == $name && $layout['masterid'] == $masterId) {
                return $layout;
            }
        }

        throw new Exception("Could not find slide layout $name in current layout pack.");
    }

    /**
     * Find specific slide layout index.
     *
     * @param string $name
     * @param int $masterId
     * @return int
     * @throws Exception
     */
    public function findLayoutIndex($name = '', $masterId = 1)
    {
        $index = 0;
        foreach ($this->_layouts as $layout) {
            if ($layout['name'] == $name && $layout['masterid'] == $masterId) {
                return $index;
            }
            ++$index;
        }

        throw new Exception("Could not find slide layout $name in current layout pack.");
    }
}
