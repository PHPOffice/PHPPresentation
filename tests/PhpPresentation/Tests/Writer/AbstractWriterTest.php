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

require 'AbstractWriter.php';

/**
 * Test class for AbstractWriter
 *
 * @coversDefaultClass AbstractWriter
 */
class AbstractWriterTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $oStubWriter = $this->getMockForAbstractClass('PhpOffice\PhpPresentation\Writer\AbstractWriter');
        $oStubZip = $this->getMockForAbstractClass('PhpOffice\Common\Adapter\Zip\ZipInterface');

        $this->assertNull($oStubWriter->getZipAdapter());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Writer\\AbstractWriter', $oStubWriter->setZipAdapter($oStubZip));
        $this->assertInstanceOf('PhpOffice\\Common\\Adapter\\Zip\\ZipInterface', $oStubWriter->getZipAdapter());
    }

    /**
     * Test all drawings method
     */
    public function testAllDrawingsIncludesMasterSlides()
    {
        $presentation = new PhpPresentation();

        $activeSlide = $presentation->getActiveSlide();
        $activeSlide->createDrawingShape();

        $masterSlides = $presentation->getAllMasterSlides();
        $masterSlide = $masterSlides[0];
        $masterSlide->createDrawingShape();

        $writer = $this->getMockForAbstractClass('PhpOffice\\PhpPresentation\\Tests\\Writer\\AbstractWriter');
        $writer->setPhpPresentation($presentation);

        $drawings = $writer->allDrawings();
        $this->assertEquals(2, count($drawings), 'Number of drawings should equal two: one from normal slide and one from master slide');
    }
}
