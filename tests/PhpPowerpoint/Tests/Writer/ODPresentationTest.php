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
use PhpOffice\PhpPowerpoint\Writer\ODPresentation;

/**
 * Test class for PhpOffice\PhpPowerpoint\Writer\ODPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\Writer\ODPresentation
 */
class ODPresentationTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $objectPrefix = 'PhpOffice\\PhpPowerpoint\\Writer\\ODPresentation\\';
        $parts = array(
            'content'  => 'Content',
            'manifest' => 'Manifest',
            'meta'     => 'Meta',
            'mimetype' => 'Mimetype',
            'styles'   => 'Styles',
            'drawing'  => 'Drawing',
        );

        $phpPowerPoint = new PhpPowerpoint();
        $phpPowerPoint->getActiveSlide()->createDrawingShape();
        $object = new ODPresentation($phpPowerPoint);

        $this->assertInstanceOf('PhpOffice\\PhpPowerpoint\\PhpPowerpoint', $object->getPHPPowerPoint());
        $this->assertEquals('./', $object->getDiskCachingDirectory());
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

        $object = new ODPresentation($phpPowerPoint);
        $object->save($filename);

        $this->assertTrue(file_exists($filename));

        unlink($filename);
    }

    /**
     * Test get writer part null
     */
    public function testGetWriterPartNull()
    {
        $object = new ODPresentation(new PhpPowerpoint());

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
        $object = new ODPresentation();
        $object->getPHPPowerPoint();
    }

    /**
     * Test set/get disk caching
     */
    public function testSetGetUseDiskCaching()
    {
        $object = new ODPresentation(new PhpPowerpoint());
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
        $object = new ODPresentation(new PhpPowerpoint());
        $object->setUseDiskCaching(true, 'foo');
    }
}
