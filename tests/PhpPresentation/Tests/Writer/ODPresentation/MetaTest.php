<?php

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

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
            ->setLastModifiedBy('ÊtaModifier')
        ;
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
}
