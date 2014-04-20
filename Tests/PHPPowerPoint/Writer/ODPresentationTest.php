<?php
/**
 * PHPPowerPoint
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2014 PHPPowerPoint
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt LGPL
 */

/**
 * Test class for PHPPowerPoint_Writer_ODPresentation
 *
 * @coversDefaultClass PHPPowerPoint_Writer_ODPresentation
 */
class PHPPowerPoint_Writer_ODPresentationTest extends PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $objectPrefix = 'PHPPowerPoint_Writer_ODPresentation_';
        $parts = array(
            'content'  => 'Content',
            'manifest' => 'Manifest',
            'meta'     => 'Meta',
            'mimetype' => 'Mimetype',
            'styles'   => 'Styles',
            'drawing'  => 'Drawing',
        );

        $phpPowerPoint = new PHPPowerPoint();
        $phpPowerPoint->getActiveSlide()->createDrawingShape();
        $object = new PHPPowerPoint_Writer_ODPresentation($phpPowerPoint);

        $this->assertInstanceOf('PHPPowerPoint', $object->getPHPPowerPoint());
        $this->assertEquals('./', $object->getDiskCachingDirectory());
        foreach ($parts as $partName => $objectName) {
            $this->assertInstanceOf($objectPrefix . $objectName, $object->getWriterPart($partName));
        }
        $this->assertInstanceOf('PHPPowerPoint_HashTable', $object->getDrawingHashTable());
    }

    /**
     * Test save
     */
    public function testSave()
    {
        $filename = tempnam(sys_get_temp_dir(), 'PHPPowerPoint');
        $imageFile = dirname(__FILE__) . '/../../resources/images/PHPPowerPointLogo.png';

        $phpPowerPoint = new PHPPowerPoint();
        $slide = $phpPowerPoint->getActiveSlide();
        $slide->createRichTextShape();
        $slide->createLineShape(10, 10, 10, 10);
        $slide->createChartShape()->getPlotArea()->setType(new PHPPowerPoint_Shape_Chart_Type_Bar3D());
        $slide->createDrawingShape()->setName('Drawing')->setPath($imageFile);
        $slide->createTableShape()->createRow();

        $object = new PHPPowerPoint_Writer_ODPresentation($phpPowerPoint);
        $object->save($filename);

        $this->assertTrue(file_exists($filename));

        unlink($filename);
    }

    /**
     * Test get writer part null
     */
    public function testGetWriterPartNull()
    {
        $object = new PHPPowerPoint_Writer_ODPresentation(new PHPPowerPoint());

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
        $object = new PHPPowerPoint_Writer_ODPresentation();
        $object->getPHPPowerPoint();
    }

    /**
     * Test set/get disk caching
     */
    public function testSetGetUseDiskCaching()
    {
        $object = new PHPPowerPoint_Writer_ODPresentation(new PHPPowerPoint());
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
        $object = new PHPPowerPoint_Writer_ODPresentation(new PHPPowerPoint());
        $object->setUseDiskCaching(true, 'foo');
    }
}
