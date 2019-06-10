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

use PhpOffice\PhpPresentation\Shape\Placeholder;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Table element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\Table
 */
class PlaceholderTest extends TestCase
{
    public function testConstruct()
    {
        $object = new Placeholder(Placeholder::PH_TYPE_BODY);
        $this->assertEquals(Placeholder::PH_TYPE_BODY, $object->getType());
        $this->assertNull($object->getIdx());
    }

    public function testIdx()
    {
        $value = mt_rand(0, 100);

        $object = new Placeholder(Placeholder::PH_TYPE_BODY);
        $this->assertNull($object->getIdx());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Placeholder', $object->setIdx($value));
        $this->assertEquals($value, $object->getIdx());
    }

    public function testType()
    {
        $rcPlaceholder = new \ReflectionClass('PhpOffice\PhpPresentation\Shape\Placeholder');
        $arrayConstants = $rcPlaceholder->getConstants();
        $value = array_rand($arrayConstants);

        $object = new Placeholder(Placeholder::PH_TYPE_BODY);
        $this->assertEquals(Placeholder::PH_TYPE_BODY, $object->getType());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Placeholder', $object->setType($value));
        $this->assertEquals($value, $object->getType());
    }
}
