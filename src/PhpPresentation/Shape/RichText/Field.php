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

namespace PhpOffice\PhpPresentation\Shape\RichText;

/**
 * Rich text field: a run whose text the application recomputes, such as a slide number or a date.
 *
 * The text it carries is the value the field last had, shown by whoever cannot compute it.
 */
class Field extends Run
{
    /**
     * The number of the slide the field sits on.
     */
    public const TYPE_SLIDENUM = 'slidenum';

    /**
     * The number of slides in the presentation.
     */
    public const TYPE_SLIDECOUNT = 'slidecount';

    /**
     * The date the presentation is shown.
     */
    public const TYPE_DATETIME = 'datetime';

    /**
     * @var string
     */
    private $type;

    public function __construct(string $type = self::TYPE_SLIDENUM, string $pText = '')
    {
        parent::__construct($pText);

        $this->type = $type;
    }

    public function getType(): string
    {
        return $this->type;
    }

    /**
     * The type is an open string rather than a list: the names above are the ones every
     * application knows, and one it does not know it shows as the text the field last had.
     */
    public function setType(string $type): self
    {
        $this->type = $type;

        return $this;
    }

    /**
     * Get hash code.
     */
    public function getHashCode(): string
    {
        return md5($this->type . parent::getHashCode() . __CLASS__);
    }
}
