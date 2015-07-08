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
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
use PhpOffice\PhpPresentation\Writer\ODPresentation;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
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
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = $oSlide->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR.'/resources/images/PhpPresentationLogo.png');
        
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $element = '/manifest:manifest/manifest:file-entry[5]';
        $this->assertTrue($pres->elementExists($element, 'META-INF/manifest.xml'));
        $this->assertEquals('Pictures/'.md5($oShape->getPath()).'.'.$oShape->getExtension(), $pres->getElementAttribute($element, 'manifest:full-path', 'META-INF/manifest.xml'));
    }
    
    public function testMemoryDrawing()
    {
        $oPhpPresentation = new PhpPresentation();
        $oSlide = $oPhpPresentation->getActiveSlide();
        $oShape = new MemoryDrawing();
        
        $gdImage = @imagecreatetruecolor(140, 20);
        $textColor = imagecolorallocate($gdImage, 255, 255, 255);
        imagestring($gdImage, 1, 5, 5, 'Created with PhpPresentation', $textColor);
        $oShape->setImageResource($gdImage)->setRenderingFunction(MemoryDrawing::RENDERING_JPEG)->setMimeType(MemoryDrawing::MIMETYPE_DEFAULT);
        $oSlide->addShape($oShape);
         
        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'ODPresentation');
        $element = '/manifest:manifest/manifest:file-entry[5]';
        $this->assertTrue($pres->elementExists($element, 'META-INF/manifest.xml'));
        $this->assertEquals('Pictures/'.$oShape->getIndexedFilename(), $pres->getElementAttribute($element, 'manifest:full-path', 'META-INF/manifest.xml'));
    }
}
