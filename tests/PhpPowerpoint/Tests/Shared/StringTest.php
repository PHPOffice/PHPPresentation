<?php
/**
 * This String is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * String that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Shared;

use PhpOffice\PhpPowerpoint\Shared\String;

/**
 * Test class for String
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shared\String
 */
class StringTest extends \PHPUnit_Framework_TestCase
{
    /**
     */
    public function testControlCharacters()
    {
        $this->assertEquals('', String::controlCharacterPHP2OOXML());
        $this->assertEquals('aeiou', String::controlCharacterPHP2OOXML('aeiou'));
        $this->assertEquals('àéîöù', String::controlCharacterPHP2OOXML('àéîöù'));

        $value = rand(0, 8);
        $this->assertEquals('_x'.sprintf('%04s', strtoupper(dechex($value))).'_', String::controlCharacterPHP2OOXML(chr($value)));
    }
    
    public function testNumberFormat()
    {
        $this->assertEquals('2.1', String::numberFormat('2.06', 1));
        $this->assertEquals('2.1', String::numberFormat('2.12', 1));
        $this->assertEquals('1234', String::numberFormat(1234, 1));
    }
}
