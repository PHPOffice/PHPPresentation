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

namespace PhpOffice\PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\Drawing;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\Rels;
use PhpOffice\PhpPresentation\Writer\ODPresentation;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation
 */
class AbstractPartTest extends \PHPUnit_Framework_TestCase
{
    protected function runProtectedMethod($obj, $method, $args = array())
    {
        $method = new \ReflectionMethod(get_class($obj), $method);
        $method->setAccessible(true);
        return $method->invokeArgs($obj, $args);
    }
    
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
    }
    
    /**
     * @expectedException           \Exception
     * @expectedExceptionMessage    No parent \PhpOffice\PhpPresentation\Writer\WriterInterface assigned.
     */
    public function testParentException()
    {
        $oDrawing = new Drawing();
        $oDrawing->getParentWriter();
    }
    
    /**
     * @expectedException           \Exception
     * @expectedExceptionMessage    The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\PowerPoint2007
     */
    public function testWriterException()
    {
        $oManifest = new Rels();
        $oManifest->setParentWriter(new ODPresentation());
        $oManifest->writeRelationships();
    }
    
    public function testXMLWriterWithDiskCaching()
    {
        $oPowerPoint2007 = new PowerPoint2007();
        $oPowerPoint2007->setUseDiskCaching(true);
        $oRels = new Rels();
        $oRels->setParentWriter($oPowerPoint2007);

        $this->assertNotEmpty(\PHPUnit_Framework_Assert::readAttribute($this->runProtectedMethod($oRels, 'getXMLWriter'), 'tempFileName'));
    }
}
