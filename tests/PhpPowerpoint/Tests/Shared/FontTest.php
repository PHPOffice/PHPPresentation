<?php
/**
 * This Font is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * Font that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Shared;

use PhpOffice\PhpPowerpoint\Shared\Font;

/**
 * Test class for Font
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Shared\Font
 */
class FontTest extends \PHPUnit_Framework_TestCase
{
    /**
     */
    public function testMath()
    {
        $value = rand(1, 100);
        $this->assertEquals(16, Font::fontSizeToPixels());
        $this->assertEquals((16 / 12) * $value, Font::fontSizeToPixels($value));
        $this->assertEquals(96, Font::inchSizeToPixels());
        $this->assertEquals(96 * $value, Font::inchSizeToPixels($value));
        $this->assertEquals(37.795275591, Font::centimeterSizeToPixels());
        $this->assertEquals(37.795275591 * $value, Font::centimeterSizeToPixels($value));
    }
}
