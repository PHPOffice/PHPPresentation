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
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Axis element.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\Chart\Axis
 */
class AxisTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Axis();

        $this->assertEquals('Axis Title', $object->getTitle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
        $this->assertNull($object->getMinorGridlines());
        $this->assertNull($object->getMajorGridlines());
    }

    public function testBounds(): void
    {
        $value = mt_rand(0, 100);
        $object = new Axis();

        $this->assertNull($object->getMinBounds());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMinBounds($value));
        $this->assertEquals($value, $object->getMinBounds());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMinBounds());
        $this->assertNull($object->getMinBounds());

        $this->assertNull($object->getMaxBounds());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMaxBounds($value));
        $this->assertEquals($value, $object->getMaxBounds());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMaxBounds());
        $this->assertNull($object->getMaxBounds());
    }

    public function testFont(): void
    {
        $object = new Axis();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setFont());
        $this->assertNull($object->getFont());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setFont(new Font()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
    }

    public function testFormatCode(): void
    {
        $object = new Axis();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setFormatCode());
        $this->assertEquals('', $object->getFormatCode());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setFormatCode('AAAA'));
        $this->assertEquals('AAAA', $object->getFormatCode());
    }

    public function testGridLines(): void
    {
        $object = new Axis();

        /** @var Gridlines $oMock */
        $oMock = $this->getMockBuilder(Gridlines::class)->getMock();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMajorGridlines($oMock));
        $this->assertInstanceOf(Gridlines::class, $object->getMajorGridlines());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMinorGridlines($oMock));
        $this->assertInstanceOf(Gridlines::class, $object->getMinorGridlines());
    }

    public function testHashIndex(): void
    {
        $object = new Axis();
        $value = mt_rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testIsVisible(): void
    {
        $object = new Axis();
        $this->assertTrue($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setIsVisible(false));
        $this->assertFalse($object->isVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setIsVisible(true));
        $this->assertTrue($object->isVisible());
    }

    public function testOutline(): void
    {
        /** @var Outline $oMock */
        $oMock = $this->getMockBuilder(Outline::class)->getMock();

        $object = new Axis();
        $this->assertInstanceOf(Outline::class, $object->getOutline());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setOutline($oMock));
        $this->assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testTickMark(): void
    {
        $value = Axis::TICK_MARK_INSIDE;
        $object = new Axis();

        $this->assertEquals(Axis::TICK_MARK_NONE, $object->getMinorTickMark());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMinorTickMark($value));
        $this->assertEquals($value, $object->getMinorTickMark());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMinorTickMark());
        $this->assertEquals(Axis::TICK_MARK_NONE, $object->getMinorTickMark());

        $this->assertEquals(Axis::TICK_MARK_NONE, $object->getMajorTickMark());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMajorTickMark($value));
        $this->assertEquals($value, $object->getMajorTickMark());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMajorTickMark());
        $this->assertEquals(Axis::TICK_MARK_NONE, $object->getMajorTickMark());
    }

    public function testTitle(): void
    {
        $object = new Axis();
        $this->assertEquals('Axis Title', $object->getTitle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setTitle('AAAA'));
        $this->assertEquals('AAAA', $object->getTitle());
    }

    public function testUnit(): void
    {
        $value = mt_rand(0, 100);
        $object = new Axis();

        $this->assertNull($object->getMinorUnit());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMinorUnit($value));
        $this->assertEquals($value, $object->getMinorUnit());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMinorUnit());
        $this->assertNull($object->getMinorUnit());

        $this->assertNull($object->getMajorUnit());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMajorUnit($value));
        $this->assertEquals($value, $object->getMajorUnit());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Axis', $object->setMajorUnit());
        $this->assertNull($object->getMajorUnit());
    }
}
