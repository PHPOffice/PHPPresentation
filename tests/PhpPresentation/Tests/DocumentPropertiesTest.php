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

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\DocumentProperties;
use PHPUnit\Framework\TestCase;

/**
 * Test class for DocumentProperties.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\DocumentProperties
 */
class DocumentPropertiesTest extends TestCase
{
    /**
     * Test get set value.
     */
    public function testGetSet(): void
    {
        $object = new DocumentProperties();
        $properties = [
            'creator' => '',
            'lastModifiedBy' => '',
            'created' => '',
            'modified' => '',
            'title' => '',
            'description' => '',
            'subject' => '',
            'keywords' => '',
            'category' => '',
            'company' => '',
        ];

        foreach ($properties as $key => $val) {
            $get = "get{$key}";
            $set = "set{$key}";
            $object->$set($val);
            $this->assertEquals($val, $object->$get());
        }
    }

    /**
     * Test get set with null value.
     */
    public function testGetSetNull(): void
    {
        $object = new DocumentProperties();
        $properties = [
            'created' => '',
            'modified' => '',
        ];
        $time = time();

        foreach (array_keys($properties) as $key) {
            $get = "get{$key}";
            $set = "set{$key}";
            $object->$set();
            $this->assertEquals($time, $object->$get());
        }
    }

    public function testCustomProperties(): void
    {
        $valueTime = time();

        $object = new DocumentProperties();
        $this->assertIsArray($object->getCustomProperties());
        $this->assertCount(0, $object->getCustomProperties());
        $this->assertFalse($object->isCustomPropertySet('pName'));
        $this->assertNull($object->getCustomPropertyType('pName'));
        $this->assertNull($object->getCustomPropertyValue('pName'));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->setCustomProperty('pName', 'pValue', null));
        $this->assertCount(1, $object->getCustomProperties());
        $this->assertTrue($object->isCustomPropertySet('pName'));
        $this->assertEquals(DocumentProperties::PROPERTY_TYPE_STRING, $object->getCustomPropertyType('pName'));
        $this->assertEquals('pValue', $object->getCustomPropertyValue('pName'));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->setCustomProperty('pName', 2, null));
        $this->assertCount(1, $object->getCustomProperties());
        $this->assertTrue($object->isCustomPropertySet('pName'));
        $this->assertEquals(DocumentProperties::PROPERTY_TYPE_INTEGER, $object->getCustomPropertyType('pName'));
        $this->assertEquals(2, $object->getCustomPropertyValue('pName'));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->setCustomProperty('pName', 2.1, null));
        $this->assertCount(1, $object->getCustomProperties());
        $this->assertTrue($object->isCustomPropertySet('pName'));
        $this->assertEquals(DocumentProperties::PROPERTY_TYPE_FLOAT, $object->getCustomPropertyType('pName'));
        $this->assertEquals(2.1, $object->getCustomPropertyValue('pName'));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->setCustomProperty('pName', true, null));
        $this->assertCount(1, $object->getCustomProperties());
        $this->assertTrue($object->isCustomPropertySet('pName'));
        $this->assertEquals(DocumentProperties::PROPERTY_TYPE_BOOLEAN, $object->getCustomPropertyType('pName'));
        $this->assertEquals(true, $object->getCustomPropertyValue('pName'));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->setCustomProperty('pName', null, null));
        $this->assertCount(1, $object->getCustomProperties());
        $this->assertTrue($object->isCustomPropertySet('pName'));
        $this->assertEquals(DocumentProperties::PROPERTY_TYPE_STRING, $object->getCustomPropertyType('pName'));
        $this->assertEquals(null, $object->getCustomPropertyValue('pName'));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->setCustomProperty('pName', $valueTime, DocumentProperties::PROPERTY_TYPE_DATE));
        $this->assertCount(1, $object->getCustomProperties());
        $this->assertTrue($object->isCustomPropertySet('pName'));
        $this->assertEquals(DocumentProperties::PROPERTY_TYPE_DATE, $object->getCustomPropertyType('pName'));
        $this->assertEquals($valueTime, $object->getCustomPropertyValue('pName'));

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\DocumentProperties', $object->setCustomProperty('pName', (string) $valueTime, DocumentProperties::PROPERTY_TYPE_UNKNOWN));
        $this->assertCount(1, $object->getCustomProperties());
        $this->assertTrue($object->isCustomPropertySet('pName'));
        $this->assertEquals(DocumentProperties::PROPERTY_TYPE_STRING, $object->getCustomPropertyType('pName'));
        $this->assertEquals($valueTime, $object->getCustomPropertyValue('pName'));
    }
}
