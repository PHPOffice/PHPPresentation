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
 * Rich text run.
 */
class Run extends TextElement implements TextElementInterface
{
    /**
     * Font.
     *
     * @var Font
     */
    private $font;

    /**
     * @param string $pText Text
     */
    public function __construct($pText = '')
    {
        // Initialize variables
        $this->setText($pText);
        $this->font = new Font();
    }

    /**
     * Get font.
     */
    public function getFont(): Font
    {
        return $this->font;
    }

    /**
     * Set font.
     *
     * @return self
     */
    public function setFont(?Font $pFont = null)
    {
        $this->font = $pFont;

        return $this;
    }

    /**
     * Get hash code.
     */
    public function getHashCode(): string
    {
        return md5($this->getText() . $this->font->getHashCode() . __CLASS__);
    }
}
