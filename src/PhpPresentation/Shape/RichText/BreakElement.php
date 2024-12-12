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

use PhpOffice\PhpPresentation\Style\Font;

/**
 * Rich text break.
 */
class BreakElement implements TextElementInterface
{
    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\RichText\Break instance.
     */
    public function __construct()
    {
    }

    /**
     * Get text.
     *
     * @return string Text
     */
    public function getText()
    {
        return "\r\n";
    }

    /**
     * Set text.
     *
     * @param string $pText Text value
     */
    public function setText($pText = ''): self
    {
        return $this;
    }

    /**
     * Get font.
     */
    public function getFont(): ?Font
    {
        return null;
    }

    /**
     * Set language.
     *
     * @param string $lang
     */
    public function setLanguage($lang): self
    {
        return $this;
    }

    /**
     * Get language.
     */
    public function getLanguage(): ?string
    {
        return null;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5(__CLASS__);
    }
}
