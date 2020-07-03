<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;

class DocPropsCustom extends AbstractDecoratorWriter
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

        // Properties
        $objWriter->startElement('Properties');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties');
        $objWriter->writeAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

        if ($this->getPresentation()->getPresentationProperties()->isMarkedAsFinal()) {
            // property
            $objWriter->startElement('property');
            $objWriter->writeAttribute('fmtid', '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}');
            $objWriter->writeAttribute('pid', 2);
            $objWriter->writeAttribute('name', '_MarkAsFinal');

            // property > vt:bool
            $objWriter->writeElement('vt:bool', 'true');

            // > property
            $objWriter->endElement();
        }

        // > Properties
        $objWriter->endElement();

        $this->oZip->addFromString('docProps/custom.xml', $objWriter->getData());

        // Return
        return $this->oZip;
    }
}
