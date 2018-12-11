<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\XMLWriter;

class Meta extends AbstractDecoratorWriter
{
    /**
     * @return ZipInterface
     */
    public function render()
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        $objWriter->startDocument('1.0', 'UTF-8');

        // office:document-meta
        $objWriter->startElement('office:document-meta');
        $objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
        $objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
        $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $objWriter->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
        $objWriter->writeAttribute('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
        $objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
        $objWriter->writeAttribute('xmlns:smil', 'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0');
        $objWriter->writeAttribute('xmlns:anim', 'urn:oasis:names:tc:opendocument:xmlns:animation:1.0');
        $objWriter->writeAttribute('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#');
        $objWriter->writeAttribute('xmlns:officeooo', 'http://openoffice.org/2009/office');
        $objWriter->writeAttribute('xmlns:drawooo', 'http://openoffice.org/2010/draw');
        $objWriter->writeAttribute('office:version', '1.2');

        // office:meta
        $objWriter->startElement('office:meta');

        // dc:creator
        $objWriter->writeElement('dc:creator', $this->getPresentation()->getDocumentProperties()->getLastModifiedBy());
        // dc:date
        $objWriter->writeElement('dc:date', gmdate('Y-m-d\TH:i:s.000', $this->getPresentation()->getDocumentProperties()->getModified()));
        // dc:description
        $objWriter->writeElement('dc:description', $this->getPresentation()->getDocumentProperties()->getDescription());
        // dc:subject
        $objWriter->writeElement('dc:subject', $this->getPresentation()->getDocumentProperties()->getSubject());
        // dc:title
        $objWriter->writeElement('dc:title', $this->getPresentation()->getDocumentProperties()->getTitle());
        // meta:creation-date
        $objWriter->writeElement('meta:creation-date', gmdate('Y-m-d\TH:i:s.000', $this->getPresentation()->getDocumentProperties()->getCreated()));
        // meta:initial-creator
        $objWriter->writeElement('meta:initial-creator', $this->getPresentation()->getDocumentProperties()->getCreator());
        // meta:keyword
        $objWriter->writeElement('meta:keyword', $this->getPresentation()->getDocumentProperties()->getKeywords());

        // @todo : Where these properties are written ?
        // $this->getPresentation()->getDocumentProperties()->getCategory()
        // $this->getPresentation()->getDocumentProperties()->getCompany()

        $objWriter->endElement();

        $objWriter->endElement();
        
        $this->getZip()->addFromString('meta.xml', $objWriter->getData());
        return $this->getZip();
    }
}
