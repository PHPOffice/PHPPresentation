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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack.
 *
 * @deprecated 0.7
 */
abstract class AbstractLayoutPack
{
    /**
     * Master slides.
     *
     * Structure:
     * - masterid
     * - body
     *
     * @var array<int, array{'masterid': int, 'body': string}>
     */
    protected $masterSlides = [];

    /**
     * Master slide relations.
     *
     * Structure:
     * - master id
     * - id (relation id)
     * - type
     * - contentType
     * - target (full path in OpenXML package)
     * - contents (body)
     *
     * @var array<int, array<string, string>>
     */
    protected $masterSlideRels = [];

    /**
     * Themes.
     *
     * Structure:
     * - masterid
     * - body
     *
     * @var array<int, array{'masterid': int, 'body': string}>
     */
    protected $themes = [];

    /**
     * Theme relations.
     *
     * Structure:
     * - masterid
     * - id (relation id)
     * - type
     * - contentType
     * - target (full path in OpenXML package)
     * - contents (body)
     *
     * @var array<int, array<string, string>>
     */
    protected $themeRelations = [];

    /**
     * Array of slide layouts.
     *
     * These are all an array consisting of:
     * - id (int)
     * - masterid (int)
     * - name (string)
     * - body (string)
     *
     * @var array<int, array{'id': int, 'masterid': int, 'name': string, 'body': string}>
     */
    protected $layouts = [];

    /**
     * Layout relations.
     *
     * Structure:
     * - layoutId (referencing layout id in layouts array)
     * - id (relation id)
     * - type
     * - contentType
     * - target (full path in OpenXML package)
     * - contents (body)
     *
     * @var array<int, array<string, string>>
     */
    protected $layoutRelations = [];

    /**
     * Get master slides.
     *
     * @return array<int, array{'masterid': int, 'body': string}>
     */
    public function getMasterSlides(): array
    {
        return $this->masterSlides;
    }

    /**
     * Get master slide relations.
     *
     * @return array<int, array<string, string>>
     */
    public function getMasterSlideRelations(): array
    {
        return $this->masterSlideRels;
    }

    /**
     * Get themes.
     *
     * @return array<int, array{'masterid': int, 'body': string}>
     */
    public function getThemes(): array
    {
        return $this->themes;
    }

    /**
     * Get theme relations.
     *
     * @return array<int, array<string, string>>
     */
    public function getThemeRelations(): array
    {
        return $this->themeRelations;
    }

    /**
     * Get array of slide layouts.
     *
     * @return array<int, array{'id': int, 'masterid': int, 'name': string, 'body': string}>
     */
    public function getLayouts()
    {
        return $this->layouts;
    }

    /**
     * Get array of slide layout relations.
     *
     * @return array<int, array<string, string>>
     */
    public function getLayoutRelations(): array
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
     * @return array{'id': int, 'masterid': int, 'name': string, 'body': string}
     *
     * @throws \Exception
     */
    public function findLayout(string $name = '', int $masterId = 1): array
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
     * @throws \Exception
     */
    public function findLayoutId(string $name = ''): int
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
     * @throws \Exception
     */
    public function findLayoutName(int $idLayout): string
    {
        foreach ($this->layouts as $layoutId => $layout) {
            if ($layoutId == $idLayout) {
                return $layout['name'];
            }
        }

        throw new \Exception("Could not find slide layout $idLayout in current layout pack.");
    }
}
