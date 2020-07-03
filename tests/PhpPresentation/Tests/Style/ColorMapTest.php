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

namespace PhpOffice\PhpPresentation\Tests\Style;

use PhpOffice\PhpPresentation\Style\ColorMap;

class ColorMapTest extends \PHPUnit_Framework_TestCase
{
    public function testConstruct()
    {
        $object = new ColorMap();
        $this->assertInternalType('array', $object->getMapping());
        $this->assertEquals(ColorMap::$mappingDefault, $object->getMapping());
    }

    public function testMapping()
    {
        $object = new ColorMap();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\ColorMap', $object->setMapping(array()));
        $this->assertInternalType('array', $object->getMapping());
        $this->assertCount(0, $object->getMapping());
        $array = ColorMap::$mappingDefault;
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\ColorMap', $object->setMapping($array));
        $this->assertInternalType('array', $object->getMapping());
        $this->assertEquals(ColorMap::$mappingDefault, $object->getMapping());
    }

    public function testModifier()
    {
        $object = new ColorMap();
        $key = array_rand(ColorMap::$mappingDefault);
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\ColorMap', $object->changeColor($key, 'AlphaBeta'));
        $array = $object->getMapping();
        $this->assertEquals('AlphaBeta', $array[$key]);
    }
}
