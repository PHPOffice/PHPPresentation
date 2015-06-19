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

namespace PhpOffice\PhpPowerpoint\Tests\Writer\ODPresentation;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerPoint\Shape\Group;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation\Drawing;

/**
 * Test class for PhpOffice\PhpPowerpoint\Writer\ODPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Writer\ODPresentation
 */
class DrawingTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
    }
    
    public function testGroup()
    {
        $oPhpPowerPoint = new PhpPowerpoint();
        $oSlide = $oPhpPowerPoint->getActiveSlide();
        $oGroup = $oSlide->createGroup();
        
        $oDrawing = new Drawing();
        $this->assertInternalType('array', $oDrawing->allDrawings($oPhpPowerPoint));
        $this->assertEmpty($oDrawing->allDrawings($oPhpPowerPoint));
        
        $oGroup->createDrawingShape();
        $oGroup->createDrawingShape();
        $oGroup->createDrawingShape();

        $this->assertInternalType('array', $oDrawing->allDrawings($oPhpPowerPoint));
        $this->assertCount(3, $oDrawing->allDrawings($oPhpPowerPoint));
    }
}
