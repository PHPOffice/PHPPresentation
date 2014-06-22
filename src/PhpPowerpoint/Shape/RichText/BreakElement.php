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

use PhpOffice\PhpPowerpoint\Shape\RichText\TextElementInterface;

/**
 * PHPPowerPoint_Shape_RichText_Break
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class BreakElement implements TextElementInterface
{
    /**
     * Create a new PHPPowerPoint_Shape_RichText_Break instance
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
     * @return PHPPowerPoint_Shape_RichText_TextElementInterface
     */
    public function setText($pText = '')
    {
        return $this;
    }

    /**
     * Get font
     *
     * @return PHPPowerPoint_Style_Font
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
