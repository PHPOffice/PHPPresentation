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

namespace PhpOffice\PhpPowerpoint\Shared;

/**
 * String
 */
class String
{
    /**
     * Control characters array
     *
     * @var string[]
     */
    private static $controlCharacters = array();

    /**
     * Is mbstring extension avalable?
     *
     * @var boolean
     */
    private static $isMbstringEnabled;

    /**
     * Is iconv extension avalable?
     *
     * @var boolean
     */
    private static $isIconvEnabled;

    /**
     * Build control characters array
     */
    private static function buildControlCharacters()
    {
        for ($i = 0; $i <= 19; ++$i) {
            if ($i != 9 && $i != 10 && $i != 13) {
                $find                            = '_x' . sprintf('%04s', strtoupper(dechex($i))) . '_';
                $replace                         = chr($i);
                self::$controlCharacters[$find] = $replace;
            }
        }
    }

    /**
     * Get whether mbstring extension is available
     *
     * @return boolean
     */
    public static function isMbstringEnabled()
    {
        if (isset(self::$isMbstringEnabled)) {
            return self::$isMbstringEnabled;
        }

        self::$isMbstringEnabled = function_exists('mb_convert_encoding') ? true : false;

        return self::$isMbstringEnabled;
    }

    /**
     * Get whether iconv extension is available
     *
     * @return boolean
     */
    public static function isIconvEnabled()
    {
        if (isset(self::$isIconvEnabled)) {
            return self::$isIconvEnabled;
        }

        self::$isIconvEnabled = function_exists('iconv') ? true : false;

        return self::$isIconvEnabled;
    }

    /**
     * Convert from OpenXML escaped control character to PHP control character
     *
     * Excel 2007 team:
     * ----------------
     * That's correct, control characters are stored directly in the shared-strings table.
     * We do encode characters that cannot be represented in XML using the following escape sequence:
     * _xHHHH_ where H represents a hexadecimal character in the character's value...
     * So you could end up with something like _x0008_ in a string (either in a cell value (<v>)
     * element or in the shared string <t> element.
     *
     * @param  string $value Value to unescape
     * @return string
     */
    public static function controlCharacterOOXML2PHP($value = '')
    {
        if (empty(self::$controlCharacters)) {
            self::buildControlCharacters();
        }

        return str_replace(array_keys(self::$controlCharacters), array_values(self::$controlCharacters), $value);
    }

    /**
     * Convert from PHP control character to OpenXML escaped control character
     *
     * Excel 2007 team:
     * ----------------
     * That's correct, control characters are stored directly in the shared-strings table.
     * We do encode characters that cannot be represented in XML using the following escape sequence:
     * _xHHHH_ where H represents a hexadecimal character in the character's value...
     * So you could end up with something like _x0008_ in a string (either in a cell value (<v>)
     * element or in the shared string <t> element.
     *
     * @param  string $value Value to escape
     * @return string
     */
    public static function controlCharacterPHP2OOXML($value = '')
    {
        if (empty(self::$controlCharacters)) {
            self::buildControlCharacters();
        }

        return str_replace(array_values(self::$controlCharacters), array_keys(self::$controlCharacters), $value);
    }

    /**
     * Check if a string contains UTF8 data
     *
     * @param  string  $value
     * @return boolean
     */
    public static function isUTF8($value = '')
    {
        return utf8_encode(utf8_decode($value)) === $value;
    }
}
