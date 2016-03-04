<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;

class DocPropsCore extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // cp:coreProperties
        $objWriter->startElement('cp:coreProperties');
        $objWriter->writeAttribute('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
        $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $objWriter->writeAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
        $objWriter->writeAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
        $objWriter->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

        // dc:creator
        $objWriter->writeElement('dc:creator', $this->oPresentation->getDocumentProperties()->getCreator());

        // cp:lastModifiedBy
        $objWriter->writeElement('cp:lastModifiedBy', $this->oPresentation->getDocumentProperties()->getLastModifiedBy());

        // dcterms:created
        $objWriter->startElement('dcterms:created');
        $objWriter->writeAttribute('xsi:type', 'dcterms:W3CDTF');
        $objWriter->writeRaw(gmdate("Y-m-d\TH:i:s\Z", $this->oPresentation->getDocumentProperties()->getCreated()));
        $objWriter->endElement();

        // dcterms:modified
        $objWriter->startElement('dcterms:modified');
        $objWriter->writeAttribute('xsi:type', 'dcterms:W3CDTF');
        $objWriter->writeRaw(gmdate("Y-m-d\TH:i:s\Z", $this->oPresentation->getDocumentProperties()->getModified()));
        $objWriter->endElement();

        // dc:title
        $objWriter->writeElement('dc:title', $this->oPresentation->getDocumentProperties()->getTitle());

        // dc:description
        $objWriter->writeElement('dc:description', $this->oPresentation->getDocumentProperties()->getDescription());

        // dc:subject
        $objWriter->writeElement('dc:subject', $this->oPresentation->getDocumentProperties()->getSubject());

        // cp:keywords
        $objWriter->writeElement('cp:keywords', $this->oPresentation->getDocumentProperties()->getKeywords());

        // cp:category
        $objWriter->writeElement('cp:category', $this->oPresentation->getDocumentProperties()->getCategory());

        if ($this->oPresentation->getPresentationProperties()->isMarkedAsFinal()) {
            // cp:contentStatus = Final
            $objWriter->writeElement('cp:contentStatus', 'Final');
        }

        $objWriter->endElement();

        $this->oZip->addFromString('docProps/core.xml', $objWriter->getData());

        // Return
        return $this->oZip;
    }
}
