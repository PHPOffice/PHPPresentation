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

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack
 * @deprecated 0.7
 */
abstract class AbstractLayoutPack
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
    protected $masterSlides = array();

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
    protected $masterSlideRels = array();

    /**
     * Themes
     *
     * Structure:
     * - masterid
     * - body
     *
     * @var array
     */
    protected $themes = '';

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
    protected $themeRelations = array();

    /**
     * Array of slide layouts.
     *
     * These are all an array consisting of:
     * - id (int)
     * - masterid (int)
     * - name (string)
     * - body (string)
     *
     * @var array
     */
    protected $layouts = array();

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
    protected $layoutRelations = array();

    /**
     * Get master slides
     *
     * @return array
     */
    public function getMasterSlides()
    {
        return $this->masterSlides;
    }

    /**
     * Get master slide relations
     *
     * @return array
     */
    public function getMasterSlideRelations()
    {
        return $this->masterSlideRels;
    }

    /**
     * Get themes
     *
     * @return array
     */
    public function getThemes()
    {
        return $this->themes;
    }

    /**
     * Get theme relations
     *
     * @return array
     */
    public function getThemeRelations()
    {
        return $this->themeRelations;
    }

    /**
     * Get array of slide layouts
     *
     * @return array
     */
    public function getLayouts()
    {
        return $this->layouts;
    }

    /**
     * Get array of slide layout relations
     *
     * @return array
     */
    public function getLayoutRelations()
    {
        return $this->layoutRelations;
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
     * @throws \Exception
     */
    public function findLayout($name = '', $masterId = 1)
    {
        foreach ($this->layouts as $layout) {
            if ($layout['name'] == $name && $layout['masterid'] == $masterId) {
                return $layout;
            }
        }

        throw new \Exception("Could not find slide layout $name in current layout pack.");
    }

    /**
     * Find specific slide layout id.
     *
     * @param string $name
     * @return int
     * @throws \Exception
     */
    public function findLayoutId($name = '')
    {
        foreach ($this->layouts as $layoutId => $layout) {
            if ($layout['name'] == $name) {
                return $layoutId;
            }
        }

        throw new \Exception("Could not find slide layout $name in current layout pack.");
    }

    /**
     * Find specific slide layout name.
     *
     * @param int $idLayout
     * @return int
     * @throws \Exception
     */
    public function findLayoutName($idLayout = null)
    {
        foreach ($this->layouts as $layoutId => $layout) {
            if ($layoutId == $idLayout) {
                return $layout['name'];
            }
        }

        throw new \Exception("Could not find slide layout $idLayout in current layout pack.");
    }
}
