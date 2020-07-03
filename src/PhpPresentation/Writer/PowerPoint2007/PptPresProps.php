<?php
namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;

class PptPresProps extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        $presentationPpts = $this->oPresentation->getPresentationProperties();

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:presentationPr
        $objWriter->startElement('p:presentationPr');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // p:presentationPr > p:showPr
        if ($presentationPpts->isLoopContinuouslyUntilEsc()) {
            $objWriter->startElement('p:showPr');
            $objWriter->writeAttribute('loop', '1');
            $objWriter->endElement();
        }

        // p:extLst
        $objWriter->startElement('p:extLst');

        // p:ext
        $objWriter->startElement('p:ext');
        $objWriter->writeAttribute('uri', '{E76CE94A-603C-4142-B9EB-6D1370010A27}');

        // p14:discardImageEditData
        $objWriter->startElement('p14:discardImageEditData');
        $objWriter->writeAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        // > p:ext
        $objWriter->endElement();

        // p:ext
        $objWriter->startElement('p:ext');
        $objWriter->writeAttribute('uri', '{D31A062A-798A-4329-ABDD-BBA856620510}');

        // p14:defaultImageDpi
        $objWriter->startElement('p14:defaultImageDpi');
        $objWriter->writeAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
        $objWriter->writeAttribute('val', '220');
        $objWriter->endElement();

        // > p:ext
        $objWriter->endElement();
        // > p:extLst
        $objWriter->endElement();
        // > p:presentationPr
        $objWriter->endElement();

        $this->getZip()->addFromString('ppt/presProps.xml', $objWriter->getData());

        return $this->getZip();
    }
}
