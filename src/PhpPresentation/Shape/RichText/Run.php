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

namespace PhpOffice\PhpPresentation\Shape\RichText;

use PhpOffice\PhpPresentation\Style\Font;

/**
 * Rich text run
 */
class Run extends TextElement implements TextElementInterface
{
    /**
     * Font
     *
     * @var \PhpOffice\PhpPresentation\Style\Font
     */
    private $font;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\RichText\Run instance
     *
     * @param string $pText Text
     */
    public function __construct($pText = '')
    {
        // Initialise variables
        $this->setText($pText);
        $this->font = new Font();
    }

    /**
     * Get font
     *
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param  \PhpOffice\PhpPresentation\Style\Font                  $pFont Font
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface
     */
    public function setFont(Font $pFont = null)
    {
        $this->font = $pFont;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->getText() . $this->font->getHashCode() . __CLASS__);
    }
}
