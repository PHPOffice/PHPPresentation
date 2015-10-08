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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart\Type;

use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Shape\Chart\Series;

/**
 * Test class for Scatter element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter
 */
class ScatterTest extends \PHPUnit_Framework_TestCase
{
    public function testData()
    {
        $object = new Scatter();

        $this->assertInternalType('array', $object->getData());
        $this->assertEmpty($object->getData());

        $array = array(
            new Series(),
            new Series(),
        );

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->setData());
        $this->assertEmpty($object->getData());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->setData($array));
        $this->assertCount(count($array), $object->getData());
    }

    public function testSerties()
    {
        $object = new Scatter();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Type\\Scatter', $object->addSeries(new Series()));
        $this->assertCount(1, $object->getData());
    }

    public function testHashCode()
    {
        $oSeries = new Series();

        $object = new Scatter();
        $object->addSeries($oSeries);

        $this->assertEquals(md5($oSeries->getHashCode().get_class($object)), $object->getHashCode());
    }
}
