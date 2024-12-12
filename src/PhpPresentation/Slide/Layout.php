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

namespace PhpOffice\PhpPresentation\Slide;

/**
 * \PhpOffice\PhpPresentation\Slide\Layout.
 */
class Layout
{
    /** Layout constants */
    public const TITLE_SLIDE = 'Title Slide';
    public const TITLE_AND_CONTENT = 'Title and Content';
    public const SECTION_HEADER = 'Section Header';
    public const TWO_CONTENT = 'Two Content';
    public const COMPARISON = 'Comparison';
    public const TITLE_ONLY = 'Title Only';
    public const BLANK = 'Blank';
    public const CONTENT_WITH_CAPTION = 'Content with Caption';
    public const PICTURE_WITH_CAPTION = 'Picture with Caption';
    public const TITLE_AND_VERTICAL_TEXT = 'Title and Vertical Text';
    public const VERTICAL_TITLE_AND_TEXT = 'Vertical Title and Text';
}
