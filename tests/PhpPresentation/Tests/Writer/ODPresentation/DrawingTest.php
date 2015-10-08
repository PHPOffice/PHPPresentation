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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Writer\ODPresentation\Drawing;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation
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
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oGroup = $oSlide->createGroup();
        
        $oDrawing = new Drawing();
        $this->assertInternalType('array', $oDrawing->allDrawings($oPhpPresentation));
        $this->assertEmpty($oDrawing->allDrawings($oPhpPresentation));
        
        $oGroup->createDrawingShape();
        $oGroup->createDrawingShape();
        $oGroup->createDrawingShape();

        $this->assertInternalType('array', $oDrawing->allDrawings($oPhpPresentation));
        $this->assertCount(3, $oDrawing->allDrawings($oPhpPresentation));
    }
}
