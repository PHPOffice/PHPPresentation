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

namespace PhpOffice\PhpPresentation\Shape;

class Placeholder
{
    /** Placeholder Type constants */
    const PH_TYPE_BODY = 'body';
    const PH_TYPE_CHART = 'chart';
    const PH_TYPE_SUBTITLE = 'subTitle';
    const PH_TYPE_TITLE = 'title';
    const PH_TYPE_FOOTER = 'ftr';
    const PH_TYPE_DATETIME = 'dt';
    const PH_TYPE_SLIDENUM = 'sldNum';
    /**
     * hasCustomPrompt
     * Indicates whether the placeholder should have a customer prompt.
     *
     * @var bool
     */
    protected $hasCustomPrompt;
    /**
     * idx
     * Specifies the index of the placeholder. This is used when applying templates or changing layouts to
     * match a placeholder on one template or master to another.
     *
     * @var int
     */
    protected $idx;
    /**
     * type
     * Specifies what content type the placeholder is to contains
     */
    protected $type;

    /**
     * Placeholder constructor.
     * @param $type
     */
    public function __construct($type)
    {
        $this->type = $type;
    }

    /**
     * @return mixed
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * @param mixed $type
     * @return Placeholder
     */
    public function setType($type)
    {
        $this->type = $type;
        return $this;
    }

    /**
     * @return int
     */
    public function getIdx()
    {
        return $this->idx;
    }

    /**
     * @param int $idx
     * @return Placeholder
     */
    public function setIdx($idx)
    {
        $this->idx = $idx;
        return $this;
    }
}
