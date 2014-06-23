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

/**
 * Rich text break
 */
class BreakElement implements TextElementInterface
{
    /**
     * Create a new \PhpOffice\PhpPowerpoint\Shape\RichText\Break instance
     */
    public function __construct()
    {
    }

    /**
     * Get text
     *
     * @return string Text
     */
    public function getText()
    {
        return "\r\n";
    }

    /**
     * Set text
     *
     * @param                                            $pText string   Text
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\TextElementInterface
     */
    public function setText($pText = '')
    {
        return $this;
    }

    /**
     * Get font
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function getFont()
    {
        return null;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5(__CLASS__);
    }
}
