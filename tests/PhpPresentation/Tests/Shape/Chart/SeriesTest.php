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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;

/**
 * Test class for Series element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart\Series
 */
class SeriesTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
        $this->assertEquals('Calibri', $object->getFont()->getName());
        $this->assertEquals(9, $object->getFont()->getSize());
        $this->assertEquals('Series Title', $object->getTitle());
        $this->assertInternalType('array', $object->getValues());
        $this->assertEmpty($object->getValues());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Marker', $object->getMarker());
        $this->assertNull($object->getOutline());
        $this->assertFalse($object->hasShowLegendKey());
    }
    
    public function testDataLabelNumFormat()
    {
        $object = new Series();
        
        $this->assertEmpty($object->getDlblNumFormat());
        $this->assertFalse($object->hasDlblNumFormat());
        
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setDlblNumFormat('#%'));
        
        $this->assertEquals('#%', $object->getDlblNumFormat());
        $this->assertTrue($object->hasDlblNumFormat());
        
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setDlblNumFormat());
        
        $this->assertEmpty($object->getDlblNumFormat());
        $this->assertFalse($object->hasDlblNumFormat());
    }

    public function testDataPointFills()
    {
        $object = new Series();

        $this->assertInternalType('array', $object->getDataPointFills());
        $this->assertEmpty($object->getDataPointFills());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getDataPointFill(0));
    }

    public function testFill()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setFill());
        $this->assertNull($object->getFill());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setFill(new Fill()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Fill', $object->getFill());
    }

    public function testFont()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setFont());
        $this->assertNull($object->getFont());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setFont(new Font()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Font', $object->getFont());
    }

    public function testHashIndex()
    {
        $object = new Series();
        $value = rand(1, 100);

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setHashIndex($value));
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testHashCode()
    {
        $object = new Series();

        $this->assertEquals(md5($object->getFill()->getHashCode().$object->getFont()->getHashCode().var_export($object->getValues(), true) . var_export($object, true).get_class($object)), $object->getHashCode());
    }

    public function testLabelPosition()
    {
        $object = new Series();

        $this->assertEmpty($object->getHashIndex());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setLabelPosition(Series::LABEL_INSIDEBASE));
        $this->assertEquals(Series::LABEL_INSIDEBASE, $object->getLabelPosition());
    }

    public function testMarker()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setMarker(new Marker()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Marker', $object->getMarker());
    }

    public function testOutline()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setOutline(new Outline()));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Outline', $object->getOutline());
    }

    public function testShowCategoryName()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowCategoryName(true));
        $this->assertTrue($object->hasShowCategoryName());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowCategoryName(false));
        $this->assertFalse($object->hasShowCategoryName());
    }

    public function testShowLeaderLines()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowLeaderLines(true));
        $this->assertTrue($object->hasShowLeaderLines());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowLeaderLines(false));
        $this->assertFalse($object->hasShowLeaderLines());
    }

    public function testShowLegendKey()
    {
        $object = new Series();

        $this->assertFalse($object->hasShowLegendKey());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowLegendKey(true));
        $this->assertTrue($object->hasShowLegendKey());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowLegendKey(false));
        $this->assertFalse($object->hasShowLegendKey());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowLegendKey(1));
        $this->assertTrue($object->hasShowLegendKey());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowLegendKey(0));
        $this->assertFalse($object->hasShowLegendKey());
    }

    public function testShowPercentage()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowPercentage(true));
        $this->assertTrue($object->hasShowPercentage());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowPercentage(false));
        $this->assertFalse($object->hasShowPercentage());
    }

    public function testShowSeparator()
    {
        $value = ';';
        $object = new Series();

        $this->assertFalse($object->hasShowSeparator());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setSeparator($value));
        $this->assertEquals($value, $object->getSeparator());
        $this->assertTrue($object->hasShowSeparator());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setSeparator(''));
        $this->assertFalse($object->hasShowPercentage());
    }

    public function testShowSeriesName()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowSeriesName(true));
        $this->assertTrue($object->hasShowSeriesName());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowSeriesName(false));
        $this->assertFalse($object->hasShowSeriesName());
    }

    public function testShowValue()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowValue(true));
        $this->assertTrue($object->hasShowValue());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setShowValue(false));
        $this->assertFalse($object->hasShowValue());
    }

    public function testTitle()
    {
        $object = new Series();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setTitle());
        $this->assertEquals('Series Title', $object->getTitle());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setTitle('AAAA'));
        $this->assertEquals('AAAA', $object->getTitle());
    }

    public function testValue()
    {
        $object = new Series();

        $array = array(
            '0' => 'a',
            '1' => 'b',
            '2' => 'c',
            '3' => 'd',
        );

        $this->assertInternalType('array', $object->getValues());
        $this->assertEmpty($object->getValues());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setValues());
        $this->assertEmpty($object->getValues());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->setValues($array));
        $this->assertCount(count($array), $object->getValues());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $object->addValue(4, 'e'));
        $this->assertCount(count($array) + 1, $object->getValues());
    }

    public function testClone()
    {
        $object = new Series();
        $object->setOutline(new Outline());
        $clone = clone $object;

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Series', $clone);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Outline', $clone->getOutline());
    }
}
