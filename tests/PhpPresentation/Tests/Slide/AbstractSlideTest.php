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

namespace PhpOffice\PhpPresentation\Tests\Slide;

use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide\AbstractSlide;

/**
 * Test class for Table element
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Shape\AbstractGraphic
 */
class AbstractSlideTest extends \PHPUnit_Framework_TestCase
{
    public function testCollection()
    {
        /** @var AbstractSlide $stub */
        $stub = $this->getMockForAbstractClass('PhpOffice\\PhpPresentation\\Slide\\AbstractSlide');

        $array = array();
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\AbstractSlide', $stub->setShapeCollection($array));
        $this->assertInternalType('array', $stub->getShapeCollection());
        $this->assertCount(count($array), $stub->getShapeCollection());

        $array = array(
            new RichText(),
            new RichText(),
            new RichText(),
        );
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\AbstractSlide', $stub->setShapeCollection($array));
        $this->assertInternalType('array', $stub->getShapeCollection());
        $this->assertCount(count($array), $stub->getShapeCollection());
    }
}
