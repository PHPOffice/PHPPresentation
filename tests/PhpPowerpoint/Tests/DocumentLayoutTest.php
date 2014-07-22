<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

namespace PhpOffice\PhpPowerpoint\Tests;

use PhpOffice\PhpPowerpoint\DocumentLayout;

/**
 * Test class for DocumentLayout
 *
 * @coversDefaultClass PhpOffice\PhpPowerpoint\DocumentLayout
 */
class DocumentLayoutTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $object = new DocumentLayout();

        $this->assertEquals('screen4x3', $object->getDocumentLayout());
        $this->assertEquals(9144000, $object->getCX());
        $this->assertEquals(6858000, $object->getCY());
        $this->assertEquals(9144000 / 36000, $object->getLayoutXmilli());
        $this->assertEquals(6858000 / 36000, $object->getLayoutYmilli());
    }

    /**
     * Test set custom layout
     */
    public function testSetCustomLayout()
    {
        $object = new DocumentLayout();
        $object->setDocumentLayout(array('cx' => 6858000, 'cy' => 9144000), false);
        $object->setLayoutXmilli(6858000 / 36000);
        $object->setLayoutYmilli(9144000 / 36000);

        $this->assertEquals('', $object->getDocumentLayout());
        $this->assertEquals(6858000, $object->getCX());
        $this->assertEquals(9144000, $object->getCY());
    }
}
