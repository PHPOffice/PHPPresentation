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
        $this->assertInstanceOf(Font::class, $object->getFont());
        $this->assertNull($object->getMinorGridlines());
        $this->assertNull($object->getMajorGridlines());
    }

    public function testBounds(): void
    {
        $value = mt_rand(0, 100);
        $object = new Axis();

        $this->assertNull($object->getMinBounds());
        $this->assertInstanceOf(Axis::class, $object->setMinBounds($value));
        $this->assertEquals($value, $object->getMinBounds());
        $this->assertInstanceOf(Axis::class, $object->setMinBounds());
        $this->assertNull($object->getMinBounds());

        $this->assertNull($object->getMaxBounds());
        $this->assertInstanceOf(Axis::class, $object->setMaxBounds($value));
        $this->assertEquals($value, $object->getMaxBounds());
        $this->assertInstanceOf(Axis::class, $object->setMaxBounds());
        $this->assertNull($object->getMaxBounds());
    }

    public function testCrossesAt(): void
    {
        $object = new Axis();

        $this->assertEquals(Axis::CROSSES_AUTO, $object->getCrossesAt());
        $this->assertInstanceOf(Axis::class, $object->setCrossesAt(Axis::CROSSES_MAX));
        $this->assertEquals(Axis::CROSSES_MAX, $object->getCrossesAt());
    }

    public function testIsReversedOrder(): void
    {
        $object = new Axis();
        $this->assertFalse($object->isReversedOrder());
        $this->assertInstanceOf(Axis::class, $object->setIsReversedOrder(true));
        $this->assertTrue($object->isReversedOrder());
        $this->assertInstanceOf(Axis::class, $object->setIsReversedOrder(false));
        $this->assertFalse($object->isReversedOrder());
    }

    public function testFont(): void
    {
        $object = new Axis();

        $this->assertInstanceOf(Axis::class, $object->setFont());
        $this->assertNull($object->getFont());
        $this->assertInstanceOf(Axis::class, $object->setFont(new Font()));
        $this->assertInstanceOf(Font::class, $object->getFont());
    }

    public function testFormatCode(): void
    {
        $object = new Axis();
        $this->assertInstanceOf(Axis::class, $object->setFormatCode());
        $this->assertEquals('', $object->getFormatCode());
        $this->assertInstanceOf(Axis::class, $object->setFormatCode('AAAA'));
        $this->assertEquals('AAAA', $object->getFormatCode());
    }

    public function testGridLines(): void
    {
        $object = new Axis();

        /** @var Gridlines $oMock */
        $oMock = $this->getMockBuilder(Gridlines::class)->getMock();

        $this->assertInstanceOf(Axis::class, $object->setMajorGridlines($oMock));
        $this->assertInstanceOf(Gridlines::class, $object->getMajorGridlines());
        $this->assertInstanceOf(Axis::class, $object->setMinorGridlines($oMock));
        $this->assertInstanceOf(Gridlines::class, $object->getMinorGridlines());
    }

    public function testHashIndex(): void
    {
        $object = new Axis();
        $value = mt_rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf(Axis::class, $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testIsVisible(): void
    {
        $object = new Axis();
        $this->assertTrue($object->isVisible());
        $this->assertInstanceOf(Axis::class, $object->setIsVisible(false));
        $this->assertFalse($object->isVisible());
        $this->assertInstanceOf(Axis::class, $object->setIsVisible(true));
        $this->assertTrue($object->isVisible());
    }

    public function testLabelRotation(): void
    {
        $object = new Axis();
        $this->assertEquals(0, $object->getTitleRotation());
        $this->assertInstanceOf(Axis::class, $object->setTitleRotation(-1));
        $this->assertEquals(0, $object->getTitleRotation());
        $this->assertInstanceOf(Axis::class, $object->setTitleRotation(361));
        $this->assertEquals(360, $object->getTitleRotation());
        $value = rand(0, 360);
        $this->assertInstanceOf(Axis::class, $object->setTitleRotation($value));
        $this->assertEquals($value, $object->getTitleRotation());
    }

    public function testOutline(): void
    {
        /** @var Outline $oMock */
        $oMock = $this->getMockBuilder(Outline::class)->getMock();

        $object = new Axis();
        $this->assertInstanceOf(Outline::class, $object->getOutline());
        $this->assertInstanceOf(Axis::class, $object->setOutline($oMock));
        $this->assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testTickLabelPosition(): void
    {
        $object = new Axis();

        $this->assertEquals(Axis::TICK_LABEL_POSITION_NEXT_TO, $object->getTickLabelPosition());
        $this->assertInstanceOf(Axis::class, $object->setTickLabelPosition(Axis::TICK_LABEL_POSITION_HIGH));
        $this->assertEquals(Axis::TICK_LABEL_POSITION_HIGH, $object->getTickLabelPosition());
        $this->assertInstanceOf(Axis::class, $object->setTickLabelPosition(Axis::TICK_LABEL_POSITION_NEXT_TO));
        $this->assertEquals(Axis::TICK_LABEL_POSITION_NEXT_TO, $object->getTickLabelPosition());
        $this->assertInstanceOf(Axis::class, $object->setTickLabelPosition(Axis::TICK_LABEL_POSITION_LOW));
        $this->assertEquals(Axis::TICK_LABEL_POSITION_LOW, $object->getTickLabelPosition());
        $this->assertInstanceOf(Axis::class, $object->setTickLabelPosition('Unauthorized'));
        $this->assertEquals(Axis::TICK_LABEL_POSITION_LOW, $object->getTickLabelPosition());
    }

    public function testTickMark(): void
    {
        $value = Axis::TICK_MARK_INSIDE;
        $object = new Axis();

        $this->assertEquals(Axis::TICK_MARK_NONE, $object->getMinorTickMark());
        $this->assertInstanceOf(Axis::class, $object->setMinorTickMark($value));
        $this->assertEquals($value, $object->getMinorTickMark());
        $this->assertInstanceOf(Axis::class, $object->setMinorTickMark());
        $this->assertEquals(Axis::TICK_MARK_NONE, $object->getMinorTickMark());

        $this->assertEquals(Axis::TICK_MARK_NONE, $object->getMajorTickMark());
        $this->assertInstanceOf(Axis::class, $object->setMajorTickMark($value));
        $this->assertEquals($value, $object->getMajorTickMark());
        $this->assertInstanceOf(Axis::class, $object->setMajorTickMark());
        $this->assertEquals(Axis::TICK_MARK_NONE, $object->getMajorTickMark());
    }

    public function testTitle(): void
    {
        $object = new Axis();
        $this->assertEquals('Axis Title', $object->getTitle());
        $this->assertInstanceOf(Axis::class, $object->setTitle('AAAA'));
        $this->assertEquals('AAAA', $object->getTitle());
    }

    public function testUnit(): void
    {
        $value = mt_rand(0, 100);
        $object = new Axis();

        $this->assertNull($object->getMinorUnit());
        $this->assertInstanceOf(Axis::class, $object->setMinorUnit($value));
        $this->assertEquals($value, $object->getMinorUnit());
        $this->assertInstanceOf(Axis::class, $object->setMinorUnit());
        $this->assertNull($object->getMinorUnit());

        $this->assertNull($object->getMajorUnit());
        $this->assertInstanceOf(Axis::class, $object->setMajorUnit($value));
        $this->assertEquals($value, $object->getMajorUnit());
        $this->assertInstanceOf(Axis::class, $object->setMajorUnit());
        $this->assertNull($object->getMajorUnit());
    }
}
