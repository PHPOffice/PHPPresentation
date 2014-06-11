<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

namespace PhpOffice\PhpPowerpoint\Tests\Writer;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;

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
        $objectPrefix = 'PhpOffice\\PhpPowerpoint\\Writer\\PowerPoint2007\\';
        $parts = array(
            'contenttypes' => 'ContentTypes',
            'docprops'     => 'DocProps',
            'rels'         => 'Rels',
            'theme'        => 'Theme',
            'presentation' => 'Presentation',
            'slide'        => 'Slide',
            'drawing'      => 'Drawing',
            'chart'        => 'Chart',
        );

        $phpPowerPoint = new PhpPowerpoint();
        $object = new PowerPoint2007($phpPowerPoint);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\PhpPowerpoint', $object->getPhpPowerpoint());
        $this->assertEquals('./', $object->getDiskCachingDirectory());
        $this->assertInstanceOf("{$objectPrefix}LayoutPack\\PackDefault", $object->getLayoutPack());
        foreach ($parts as $partName => $objectName) {
            $this->assertInstanceOf($objectPrefix . $objectName, $object->getWriterPart($partName));
        }
        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\HashTable', $object->getDrawingHashTable());
    }

    /**
     * Test save
     */
    public function testSave()
    {
        $filename = tempnam(sys_get_temp_dir(), 'PHPPowerPoint');
        $imageFile = dirname(__FILE__) . '/../../../resources/images/PHPPowerPointLogo.png';

        $phpPowerPoint = new PhpPowerpoint();
        $slide = $phpPowerPoint->getActiveSlide();
        $slide->createRichTextShape();
        $slide->createLineShape(10, 10, 10, 10);
        $slide->createChartShape()->getPlotArea()->setType(new \PhpOffice\PhpPowerpoint\Shape\Chart\Type\Bar3D());
        $slide->createDrawingShape()->setName('Drawing')->setPath($imageFile);
        $slide->createTableShape()->createRow();

        $object = new PowerPoint2007($phpPowerPoint);
        $object->save($filename);

        $this->assertTrue(file_exists($filename));

        unlink($filename);
    }

    /**
     * Test get writer part null
     */
    public function testGetWriterPartNull()
    {
        $object = new PowerPoint2007(new PhpPowerpoint());

        $this->assertNull($object->getWriterPart('foo'));
    }

    /**
     * Test get PHPPowerPoint exception
     *
     * @expectedException Exception
     * @expectedExceptionMessage No PHPPowerPoint assigned.
     */
    public function testGetPHPPowerPointException()
    {
        $object = new PowerPoint2007();
        $object->getPhpPowerpoint();
    }

    /**
     * Test set/get Office 2003 compatibility
     */
    public function testSetGetOffice2003Compatibility()
    {
        $object = new PowerPoint2007(new PhpPowerpoint());
        $this->assertFalse($object->getOffice2003Compatibility());

        $object->setOffice2003Compatibility(true);
        $this->assertTrue($object->getOffice2003Compatibility());
    }

    /**
     * Test set/get disk caching
     */
    public function testSetGetUseDiskCaching()
    {
        $object = new PowerPoint2007(new PhpPowerpoint());
        $this->assertFalse($object->getUseDiskCaching());

        $object->setUseDiskCaching(true, sys_get_temp_dir());
        $this->assertTrue($object->getUseDiskCaching());
        $this->assertEquals(sys_get_temp_dir(), $object->getDiskCachingDirectory());
    }

    /**
     * Test set/get disk caching exception
     *
     * @expectedException Exception
     */
    public function testSetUseDiskCachingException()
    {
        $object = new PowerPoint2007(new PhpPowerpoint());
        $object->setUseDiskCaching(true, 'foo');
    }
}
