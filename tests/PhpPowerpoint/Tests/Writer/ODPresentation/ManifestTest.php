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
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation;
use PhpOffice\PhpPowerpoint\Tests\TestHelperDOCX;

/**
 * Test class for PhpOffice\PhpPowerpoint\Writer\ODPresentation\Manifest
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Writer\ODPresentation\Manifest
 */
class ManifestTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Executed before each method of the class
     */
    public function tearDown()
    {
        TestHelperDOCX::clear();
    }
    
    public function testDrawing()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPOWERPOINT_TESTS_BASE_DIR.'/resources/images/PHPPowerPointLogo.png');
        
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/manifest:manifest/manifest:file-entry[5]';
        $this->assertTrue($pres->elementExists($element, 'META-INF/manifest.xml'));
        $this->assertEquals('Pictures/'.md5($oShape->getPath()).'.'.$oShape->getExtension(), $pres->getElementAttribute($element, 'manifest:full-path', 'META-INF/manifest.xml'));
    }
    
    public function testMemoryDrawing()
    {
        $phpPowerPoint = new PhpPowerpoint();
        $oSlide = $phpPowerPoint->getActiveSlide();
        $oShape = new MemoryDrawing();
        
        $gdImage = @imagecreatetruecolor(140, 20);
        $textColor = imagecolorallocate($gdImage, 255, 255, 255);
        imagestring($gdImage, 1, 5, 5, 'Created with PHPPowerPoint', $textColor);
        $oShape->setImageResource($gdImage)->setRenderingFunction(MemoryDrawing::RENDERING_JPEG)->setMimeType(MemoryDrawing::MIMETYPE_DEFAULT);
        $oSlide->addShape($oShape);
         
        $pres = TestHelperDOCX::getDocument($phpPowerPoint, 'ODPresentation');
        $element = '/manifest:manifest/manifest:file-entry[5]';
        $this->assertTrue($pres->elementExists($element, 'META-INF/manifest.xml'));
        $this->assertEquals('Pictures/'.$oShape->getIndexedFilename(), $pres->getElementAttribute($element, 'manifest:full-path', 'META-INF/manifest.xml'));
    }
}
