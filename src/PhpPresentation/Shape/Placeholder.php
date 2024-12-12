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
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Shape;

class Placeholder
{
    /** Placeholder Type constants */
    public const PH_TYPE_BODY = 'body';
    public const PH_TYPE_CHART = 'chart';
    public const PH_TYPE_SUBTITLE = 'subTitle';
    public const PH_TYPE_TITLE = 'title';
    public const PH_TYPE_FOOTER = 'ftr';
    public const PH_TYPE_DATETIME = 'dt';
    public const PH_TYPE_SLIDENUM = 'sldNum';

    /**
     * Indicates whether the placeholder should have a customer prompt.
     *
     * @var bool
     */
    protected $hasCustomPrompt;

    /**
     * Specifies the index of the placeholder. This is used when applying templates or changing layouts to
     * match a placeholder on one template or master to another.
     *
     * @var null|int
     */
    protected $idx;

    /**
     * Specifies what content type the placeholder is to contains.
     *
     * @var string
     */
    protected $type;

    public function __construct(string $type)
    {
        $this->type = $type;
    }

    public function getType(): string
    {
        return $this->type;
    }

    public function setType(string $type): self
    {
        $this->type = $type;

        return $this;
    }

    public function getIdx(): ?int
    {
        return $this->idx;
    }

    public function setIdx(int $idx): self
    {
        $this->idx = $idx;

        return $this;
    }
}
