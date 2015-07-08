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
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
 */
class PptPropsTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }

    public function testPresProps()
    {
        $oPhpPresentation = new PhpPresentation();
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/p:presentationPr/p:extLst/p:ext';
        $this->assertTrue($pres->elementExists($element, 'ppt/presProps.xml'));
        $this->assertEquals('{E76CE94A-603C-4142-B9EB-6D1370010A27}', $pres->getElementAttribute($element, 'uri', 'ppt/presProps.xml'));
    }
    
    public function testTableStyles()
    {
        $oPhpPresentation = new PhpPresentation();
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/a:tblStyleLst';
        $this->assertTrue($pres->elementExists($element, 'ppt/tableStyles.xml'));
        $this->assertEquals('{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}', $pres->getElementAttribute($element, 'def', 'ppt/tableStyles.xml'));
    }
    
    public function testViewProps()
    {
        $oPhpPresentation = new PhpPresentation();
    
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $element = '/p:viewPr';
        $this->assertTrue($pres->elementExists($element, 'ppt/viewProps.xml'));
        $this->assertEquals('0', $pres->getElementAttribute($element, 'showComments', 'ppt/viewProps.xml'));
    }
}
