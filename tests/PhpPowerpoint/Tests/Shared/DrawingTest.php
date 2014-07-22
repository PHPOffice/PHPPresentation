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
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests\Shared;

use PhpOffice\PhpPowerpoint\Shared\Drawing;

/**
 * Test class for IOFactory
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\IOFactory
 */
class DrawingTest extends \PHPUnit_Framework_TestCase
{
    /**
     */
    public function testPixelsCentimeters()
    {
        $value = rand(1, 100);

        $this->assertEquals(0, Drawing::pixelsToCentimeters());
        $this->assertEquals($value / Drawing::DPI_96 * 2.54, Drawing::pixelsToCentimeters($value));
        $this->assertEquals(0, Drawing::centimetersToPixels());
        $this->assertEquals($value / 2.54 * Drawing::DPI_96, Drawing::centimetersToPixels($value));
    }

    /**
     */
    public function testPixelsEMU()
    {
        $value = rand(1, 100);

        $this->assertEquals(0, Drawing::pixelsToEmu());
        $this->assertEquals(round($value*9525), Drawing::pixelsToEmu($value));
        $this->assertEquals(0, Drawing::emuToPixels());
        $this->assertEquals(round($value/9525), Drawing::emuToPixels($value));
    }

    /**
     */
    public function testPixelsPoints()
    {
        $value = rand(1, 100);

        $this->assertEquals(0, Drawing::pixelsToPoints());
        $this->assertEquals($value*0.67777777, Drawing::pixelsToPoints($value));
        $this->assertEquals(0, Drawing::pointsToPixels());
        $this->assertEquals($value* 1.333333333, Drawing::pointsToPixels($value));
    }

    /**
     */
    public function testDegreesAngle()
    {
        $value = rand(1, 100);

        $this->assertEquals(0, Drawing::degreesToAngle());
        $this->assertEquals((int) round($value * 60000), Drawing::degreesToAngle($value));
        $this->assertEquals(0, Drawing::angleToDegrees());
        $this->assertEquals(round($value / 60000), Drawing::angleToDegrees($value));
    }
}
