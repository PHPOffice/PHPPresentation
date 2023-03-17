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

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class MetaTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    /**
     * @dataProvider dataProviderCustomProperties
     *
     * @param mixed $propertyValue
     * @param string|null $propertyType
     * @param string $expectedValue
     * @param string $expectedValueType
     */
    public function testCustomProperties($propertyValue, ?string $propertyType, string $expectedValue, string $expectedValueType): void
    {
        $this->oPresentation->getDocumentProperties()->setCustomProperty('pName', $propertyValue, $propertyType);

        $this->assertZipXmlElementExists('meta.xml', '/office:document-meta');
        $this->assertZipXmlElementExists('meta.xml', '/office:document-meta/office:meta');
        $this->assertZipXmlElementExists('meta.xml', '/office:document-meta/office:meta/meta:user-defined');
        $this->assertZipXmlElementEquals('meta.xml', '/office:document-meta/office:meta/meta:user-defined', $expectedValue);
        $this->assertZipXmlAttributeExists('meta.xml', '/office:document-meta/office:meta/meta:user-defined', 'meta:name');
        $this->assertZipXmlAttributeEquals('meta.xml', '/office:document-meta/office:meta/meta:user-defined', 'meta:name', 'pName');
        $this->assertZipXmlAttributeExists('meta.xml', '/office:document-meta/office:meta/meta:user-defined', 'meta:value-type');
        $this->assertZipXmlAttributeEquals('meta.xml', '/office:document-meta/office:meta/meta:user-defined', 'meta:value-type', $expectedValueType);
    }

    /**
     * @return array<array<bool|string|int|float|null>>
     */
    public function dataProviderCustomProperties(): array
    {
        $valueTime = time();

        return [
            [
                false,
                null,
                'false',
                'boolean',
            ],
            [
                $valueTime,
                DocumentProperties::PROPERTY_TYPE_DATE,
                date(DATE_W3C, $valueTime),
                'date',
            ],
            [
                2.1,
                null,
                '2.1',
                'float',
            ],
            [
                2,
                null,
                '2',
                'float',
            ],
            [
                null,
                null,
                '',
                'string',
            ],
            [
                (string) $valueTime,
                DocumentProperties::PROPERTY_TYPE_UNKNOWN,
                (string) $valueTime,
                'string',
            ],
        ];
    }
}
