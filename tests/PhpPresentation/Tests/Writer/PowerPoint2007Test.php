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

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Tests\TestHelperDOCX;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * Test class for PowerPoint2007
 *
 * @coversDefaultClass PowerPoint2007
 */
class PowerPoint2007Test extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $objectPrefix = 'PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007\\';

        $oPhpPresentation = new PhpPresentation();
        $object = new PowerPoint2007($oPhpPresentation);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PhpPresentation', $object->getPhpPresentation());
        $this->assertEquals('./', $object->getDiskCachingDirectory());
        $this->assertInstanceOf("{$objectPrefix}LayoutPack\\PackDefault", $object->getLayoutPack());
    }

    /**
     * Test save
     */
    public function testSave()
    {
        $filename = tempnam(sys_get_temp_dir(), 'PhpPresentation');

        $oPhpPresentation = new PhpPresentation();

        $object = new PowerPoint2007($oPhpPresentation);
        $object->save($filename);

        $this->assertTrue(file_exists($filename));

        unlink($filename);
    }

    /**
     * Test save with empty filename
     *
     * @expectedException Exception
     * @expectedExceptionMessage Filename is empty
     */
    public function testSaveEmptyException()
    {
        $oPhpPresentation = new PhpPresentation();

        $object = new PowerPoint2007($oPhpPresentation);
        $object->save('');
    }

    /**
     * Test save with empty assignation
     *
     * @expectedException Exception
     * @expectedExceptionMessage No PhpPresentation assigned.
     */
    public function testSaveUnassignedException()
    {
        $object = new PowerPoint2007();
        $object->save('filename.pptx');
    }

    /**
     * Test get PhpPresentation exception
     *
     * @expectedException Exception
     * @expectedExceptionMessage No PhpPresentation assigned.
     */
    public function testGetPhpPresentationException()
    {
        $object = new PowerPoint2007();
        $object->getPhpPresentation();
    }

    /**
     * Test disk caching
     */
    public function testDiskCaching()
    {
        $object = new PowerPoint2007(new PhpPresentation());
        $this->assertFalse($object->hasDiskCaching());

        $object->setUseDiskCaching(true);
        $this->assertTrue($object->hasDiskCaching());
        $this->assertEquals('./', $object->getDiskCachingDirectory());

        $object->setUseDiskCaching(true, sys_get_temp_dir());
        $this->assertTrue($object->hasDiskCaching());
        $this->assertEquals(sys_get_temp_dir(), $object->getDiskCachingDirectory());
    }

    /**
     * Test set/get disk caching exception
     *
     * @expectedException Exception
     * @expectedExceptionMessage Directory does not exist: foo
     */
    public function testCachingException()
    {
        $object = new PowerPoint2007(new PhpPresentation());
        $object->setUseDiskCaching(true, 'foo');
    }

    /**
     * Test LayoutPack
     * @deprecated 0.7
     */
    public function testLayoutPack()
    {
        $oLayoutPack = $this->getMock('PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007\\LayoutPack\\AbstractLayoutPack');

        $object = new PowerPoint2007();

        $this->assertInstanceOf("PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007\\LayoutPack\\AbstractLayoutPack", $object->getLayoutPack());
        $this->assertInstanceOf("PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007", $object->setLayoutPack());
        $this->assertNull($object->getLayoutPack());
        $this->assertInstanceOf("PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007", $object->setLayoutPack($oLayoutPack));
        $this->assertInstanceOf("PhpOffice\\PhpPresentation\\Writer\\PowerPoint2007\\LayoutPack\\AbstractLayoutPack", $object->getLayoutPack());
    }

    public function testZoom()
    {
        $oPhpPresentation = new PhpPresentation();

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->elementExists('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'ppt/viewProps.xml'));
        $this->assertEquals('100', $pres->getElementAttribute('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'n', 'ppt/viewProps.xml'));
        $this->assertEquals('100', $pres->getElementAttribute('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'd', 'ppt/viewProps.xml'));
        $this->assertTrue($pres->elementExists('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'ppt/viewProps.xml'));
        $this->assertEquals('100', $pres->getElementAttribute('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'n', 'ppt/viewProps.xml'));
        $this->assertEquals('100', $pres->getElementAttribute('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'd', 'ppt/viewProps.xml'));

        $value = rand(1, 100);
        $oPhpPresentation->getPresentationProperties()->setZoom($value);

        $pres = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($pres->elementExists('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'ppt/viewProps.xml'));
        $this->assertEquals($value * 100, $pres->getElementAttribute('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'n', 'ppt/viewProps.xml'));
        $this->assertEquals('100', $pres->getElementAttribute('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx', 'd', 'ppt/viewProps.xml'));
        $this->assertTrue($pres->elementExists('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'ppt/viewProps.xml'));
        $this->assertEquals($value * 100, $pres->getElementAttribute('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'n', 'ppt/viewProps.xml'));
        $this->assertEquals('100', $pres->getElementAttribute('/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sy', 'd', 'ppt/viewProps.xml'));
    }

    public function testFeatureThumbnail()
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';

        $xPathManifest = '/Relationships/Relationship[@Target=\'docProps/thumbnail.jpeg\'][@Type=\'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\']';
        $oPhpPresentation = new PhpPresentation();

        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('_rels/.rels'));
        $this->assertFalse($oXMLDoc->elementExists($xPathManifest, '_rels/.rels'));

        $oPhpPresentation->getPresentationProperties()->setThumbnailPath($imagePath);
        $oXMLDoc = TestHelperDOCX::getDocument($oPhpPresentation, 'PowerPoint2007');
        $this->assertTrue($oXMLDoc->fileExists('_rels/.rels'));
        $this->assertTrue($oXMLDoc->elementExists($xPathManifest, '_rels/.rels'));
    }
}
