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

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PHPUnit\Framework\Attributes\DataProvider;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Meta.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\ODPresentation\Meta
 */
class MetaTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    public function testDocumentProperties(): void
    {
        $element = '/office:document-meta/office:meta';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $element = '/office:document-meta/office:meta/dc:creator';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, 'Unknown Creator');
        $element = '/office:document-meta/office:meta/dc:date';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, gmdate('Y-m-d\TH:i:s.000', $this->oPresentation->getDocumentProperties()->getModified()));
        $element = '/office:document-meta/office:meta/dc:description';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, '');
        $element = '/office:document-meta/office:meta/dc:subject';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, '');
        $element = '/office:document-meta/office:meta/dc:title';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, 'Untitled Presentation');
        $element = '/office:document-meta/office:meta/meta:creation-date';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, gmdate('Y-m-d\TH:i:s.000', $this->oPresentation->getDocumentProperties()->getCreated()));
        $element = '/office:document-meta/office:meta/meta:initial-creator';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, 'Unknown Creator');
        $element = '/office:document-meta/office:meta/meta:keyword';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, '');
        $element = '/office:document-meta/office:meta/meta:generator';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, '');

        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->oPresentation->getDocumentProperties()
            ->setCreator('AlphaCreator')
            ->setDescription('BetaDescription')
            ->setSubject('GammaSubject')
            ->setTitle('DeltaTitle')
            ->setKeywords('EpsilonKeyword')
            ->setGenerator('ZêtaGenerator')
            ->setLastModifiedBy('ÊtaModifier');
        $this->resetPresentationFile();

        $element = '/office:document-meta/office:meta';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $element = '/office:document-meta/office:meta/dc:creator';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, $this->oPresentation->getDocumentProperties()->getLastModifiedBy());
        $element = '/office:document-meta/office:meta/dc:date';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, gmdate('Y-m-d\TH:i:s.000', $this->oPresentation->getDocumentProperties()->getModified()));
        $element = '/office:document-meta/office:meta/dc:description';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, $this->oPresentation->getDocumentProperties()->getDescription());
        $element = '/office:document-meta/office:meta/dc:subject';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, $this->oPresentation->getDocumentProperties()->getSubject());
        $element = '/office:document-meta/office:meta/dc:title';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, $this->oPresentation->getDocumentProperties()->getTitle());
        $element = '/office:document-meta/office:meta/meta:creation-date';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, gmdate('Y-m-d\TH:i:s.000', $this->oPresentation->getDocumentProperties()->getCreated()));
        $element = '/office:document-meta/office:meta/meta:initial-creator';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, $this->oPresentation->getDocumentProperties()->getCreator());
        $element = '/office:document-meta/office:meta/meta:keyword';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, $this->oPresentation->getDocumentProperties()->getKeywords());
        $element = '/office:document-meta/office:meta/meta:generator';
        $this->assertZipXmlElementExists('meta.xml', $element);
        $this->assertZipXmlElementEquals('meta.xml', $element, $this->oPresentation->getDocumentProperties()->getGenerator());
    }

    /**
     * @dataProvider dataProviderCustomProperties
     *
     * @param mixed $propertyValue
     */
    #[DataProvider('dataProviderCustomProperties')]
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
     * @return array<array<null|bool|float|int|string>>
     */
    public static function dataProviderCustomProperties(): array
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
