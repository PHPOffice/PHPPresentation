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
use PHPUnit\Framework\TestCase;

/**
 * Test class for Legend element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Chart\Marker
 */
class MarkerTest extends TestCase
{
    public function testConstruct()
    {
        $object = new Marker();

        $this->assertEquals(Marker::SYMBOL_NONE, $object->getSymbol());
        $this->assertEquals(5, $object->getSize());
    }

    public function testSymbol()
    {
        $object = new Marker();

        $expected = array_rand(Marker::$arraySymbol);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Marker', $object->setSymbol($expected));
        $this->assertEquals($expected, $object->getSymbol());
    }

    public function testSize()
    {
        $object = new Marker();
        $value = mt_rand(1, 100);

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Chart\\Marker', $object->setSize($value));
        $this->assertEquals($value, $object->getSize());
    }
}
