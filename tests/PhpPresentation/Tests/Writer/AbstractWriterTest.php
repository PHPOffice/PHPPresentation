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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Tests\Writer\AbstractWriter as TestAbstractWriter;
use PhpOffice\PhpPresentation\Writer\AbstractWriter;
use PHPUnit\Framework\TestCase;

require 'AbstractWriter.php';

/**
 * Test class for AbstractWriter.
 *
 * @coversDefaultClass \AbstractWriter
 */
class AbstractWriterTest extends TestCase
{
    /**
     * Test create new instance.
     */
    public function testConstruct(): void
    {
        /** @var AbstractWriter $oStubWriter */
        $oStubWriter = $this->getMockForAbstractClass(AbstractWriter::class);
        /** @var ZipInterface $oStubZip */
        $oStubZip = $this->getMockForAbstractClass(ZipInterface::class);

        $this->assertNull($oStubWriter->getZipAdapter());
        $this->assertInstanceOf(AbstractWriter::class, $oStubWriter->setZipAdapter($oStubZip));
        $this->assertInstanceOf(ZipInterface::class, $oStubWriter->getZipAdapter());
    }

    /**
     * Test all drawings method.
     */
    public function testAllDrawingsIncludesMasterSlides(): void
    {
        $presentation = new PhpPresentation();

        $activeSlide = $presentation->getActiveSlide();
        $activeSlide->createDrawingShape();

        $masterSlides = $presentation->getAllMasterSlides();
        $masterSlide = $masterSlides[0];
        $masterSlide->createDrawingShape();

        /** @var TestAbstractWriter $writer */
        $writer = $this->getMockForAbstractClass(TestAbstractWriter::class);
        $writer->setPhpPresentation($presentation);

        $drawings = $writer->allDrawings();
        $this->assertCount(2, $drawings, 'Number of drawings should equal two: one from normal slide and one from master slide');
    }
}
