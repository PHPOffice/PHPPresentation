<?php

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
