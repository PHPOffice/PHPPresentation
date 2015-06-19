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

namespace PhpOffice\PhpPowerpoint\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Drawing;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Rels;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation;

/**
 * Test class for PhpOffice\PhpPowerpoint\Writer\ODPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Writer\ODPresentation
 */
class AbstractPartTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
    }
    
    /**
     * @expectedException           \Exception
     * @expectedExceptionMessage    No parent \PhpOffice\PhpPowerpoint\Writer\WriterInterface assigned.
     */
    public function testParentException()
    {
        $oDrawing = new Drawing();
        $oDrawing->getParentWriter();
    }
    
    /**
     * @expectedException           \Exception
     * @expectedExceptionMessage    The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007
     */
    public function testWriterException()
    {
        $oManifest = new Rels();
        $oManifest->setParentWriter(new ODPresentation());
        $oManifest->writeRelationships();
    }
}
