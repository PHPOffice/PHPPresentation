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
            'title' => '',
            'description' => '',
            'subject' => '',
            'keywords' => '',
            'category' => '',
            'company' => '',
            'revision' => '',
            'status' => '',
            'generator' => '',
        ];

        foreach ($properties as $key => $val) {
            $get = "get{$key}";
            $set = "set{$key}";
            $object->$set($val);
            self::assertEquals($val, $object->$get());
        }
    }

    /**
     * Test get set with int value.
     */
    public function testGetSetInt(): void
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
            self::assertEquals($time, $object->$get());
        }
    }

    public function testCustomProperties(): void
    {
        $valueTime = time();

        $object = new DocumentProperties();
        self::assertIsArray($object->getCustomProperties());
        self::assertCount(0, $object->getCustomProperties());
        self::assertFalse($object->isCustomPropertySet('pName'));
        self::assertNull($object->getCustomPropertyType('pName'));
        self::assertNull($object->getCustomPropertyValue('pName'));

        self::assertInstanceOf(DocumentProperties::class, $object->setCustomProperty('pName', 'pValue', null));
        self::assertCount(1, $object->getCustomProperties());
        self::assertTrue($object->isCustomPropertySet('pName'));
        self::assertEquals(DocumentProperties::PROPERTY_TYPE_STRING, $object->getCustomPropertyType('pName'));
        self::assertEquals('pValue', $object->getCustomPropertyValue('pName'));

        self::assertInstanceOf(DocumentProperties::class, $object->setCustomProperty('pName', 2, null));
        self::assertCount(1, $object->getCustomProperties());
        self::assertTrue($object->isCustomPropertySet('pName'));
        self::assertEquals(DocumentProperties::PROPERTY_TYPE_INTEGER, $object->getCustomPropertyType('pName'));
        self::assertEquals(2, $object->getCustomPropertyValue('pName'));

        self::assertInstanceOf(DocumentProperties::class, $object->setCustomProperty('pName', 2.1, null));
        self::assertCount(1, $object->getCustomProperties());
        self::assertTrue($object->isCustomPropertySet('pName'));
        self::assertEquals(DocumentProperties::PROPERTY_TYPE_FLOAT, $object->getCustomPropertyType('pName'));
        self::assertEquals(2.1, $object->getCustomPropertyValue('pName'));

        self::assertInstanceOf(DocumentProperties::class, $object->setCustomProperty('pName', true, null));
        self::assertCount(1, $object->getCustomProperties());
        self::assertTrue($object->isCustomPropertySet('pName'));
        self::assertEquals(DocumentProperties::PROPERTY_TYPE_BOOLEAN, $object->getCustomPropertyType('pName'));
        self::assertTrue($object->getCustomPropertyValue('pName'));

        self::assertInstanceOf(DocumentProperties::class, $object->setCustomProperty('pName', null, null));
        self::assertCount(1, $object->getCustomProperties());
        self::assertTrue($object->isCustomPropertySet('pName'));
        self::assertEquals(DocumentProperties::PROPERTY_TYPE_STRING, $object->getCustomPropertyType('pName'));
        self::assertNull($object->getCustomPropertyValue('pName'));

        self::assertInstanceOf(DocumentProperties::class, $object->setCustomProperty('pName', $valueTime, DocumentProperties::PROPERTY_TYPE_DATE));
        self::assertCount(1, $object->getCustomProperties());
        self::assertTrue($object->isCustomPropertySet('pName'));
        self::assertEquals(DocumentProperties::PROPERTY_TYPE_DATE, $object->getCustomPropertyType('pName'));
        self::assertEquals($valueTime, $object->getCustomPropertyValue('pName'));

        self::assertInstanceOf(DocumentProperties::class, $object->setCustomProperty('pName', (string) $valueTime, DocumentProperties::PROPERTY_TYPE_UNKNOWN));
        self::assertCount(1, $object->getCustomProperties());
        self::assertTrue($object->isCustomPropertySet('pName'));
        self::assertEquals(DocumentProperties::PROPERTY_TYPE_STRING, $object->getCustomPropertyType('pName'));
        self::assertEquals($valueTime, $object->getCustomPropertyValue('pName'));
    }
}
