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

namespace PhpOffice\PhpPowerpoint\Shape\RichText;

use PhpOffice\PhpPowerpoint\Style\Font;

/**
 * Rich text run
 */
class Run extends TextElement implements TextElementInterface
{
    /**
     * Font
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Font
     */
    private $font;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Shape\RichText\Run instance
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
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param  \PhpOffice\PhpPowerpoint\Style\Font                  $pFont Font
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\TextElementInterface
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
