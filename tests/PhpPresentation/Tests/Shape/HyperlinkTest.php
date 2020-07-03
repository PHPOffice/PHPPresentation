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

namespace PhpOffice\PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\Hyperlink;

/**
 * Test class for hyperlink element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Hyperlink
 */
class HyperlinkTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test can read
     */
    public function testConstruct()
    {
        $object = new Hyperlink();
        $this->assertEmpty($object->getUrl());
        $this->assertEmpty($object->getTooltip());

        $object = new Hyperlink('http://test.com');
        $this->assertEquals('http://test.com', $object->getUrl());
        $this->assertEmpty($object->getTooltip());

        $object = new Hyperlink('http://test.com', 'Test');
        $this->assertEquals('http://test.com', $object->getUrl());
        $this->assertEquals('Test', $object->getTooltip());
    }

    /**
     * Test get hash code
     */
    public function testGetHashCode()
    {
        $object = new Hyperlink();
        $this->assertEquals(md5(get_class($object)), $object->getHashCode());

        $object = new Hyperlink('http://test.com');
        $this->assertEquals(md5('http://test.com'.get_class($object)), $object->getHashCode());

        $object = new Hyperlink('http://test.com', 'Test');
        $this->assertEquals(md5('http://test.com'.'Test'.get_class($object)), $object->getHashCode());
    }

    /**
     * Test get/set hash index
     */
    public function testSetGetHashIndex()
    {
        $object = new Hyperlink();
        $value = rand(1, 100);
        $object->setHashIndex($value);
        $this->assertEquals($value, $object->getHashIndex());
    }

    public function testGetSetSlideNumber()
    {
        $object = new Hyperlink();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->setSlideNumber());
        $this->assertEquals(1, $object->getSlideNumber());
        $this->assertEquals('ppaction://hlinksldjump', $object->getUrl());

        $value = rand(1, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->setSlideNumber($value));
        $this->assertEquals($value, $object->getSlideNumber());
        $this->assertEquals('ppaction://hlinksldjump', $object->getUrl());
    }

    public function testGetSetTooltip()
    {
        $object = new Hyperlink();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->setTooltip());
        $this->assertEmpty($object->getTooltip());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->setTooltip('TEST'));
        $this->assertEquals('TEST', $object->getTooltip());
    }

    public function testGetSetUrl()
    {
        $object = new Hyperlink();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->setUrl());
        $this->assertEmpty($object->getUrl());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->setUrl('http://www.github.com'));
        $this->assertEquals('http://www.github.com', $object->getUrl());
    }

    public function testIsInternal()
    {
        $object = new Hyperlink();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->setSlideNumber());
        $this->assertTrue($object->isInternal());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Hyperlink', $object->setUrl('http://www.github.com'));
        $this->assertFalse($object->isInternal());
    }
}
