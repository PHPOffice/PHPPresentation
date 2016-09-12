<?php
namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;

class PptViewProps extends AbstractDecoratorWriter
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

        // p:viewPr
        $objWriter->startElement('p:viewPr');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $objWriter->writeAttribute('showComments', $this->getPresentation()->getPresentationProperties()->isCommentVisible() ? 1 : 0);
        $objWriter->writeAttribute('lastView', $this->getPresentation()->getPresentationProperties()->getLastView());

        // p:viewPr > p:slideViewPr
        $objWriter->startElement('p:slideViewPr');

        // p:viewPr > p:slideViewPr > p:cSldViewPr
        $objWriter->startElement('p:cSldViewPr');

        // p:viewPr > p:slideViewPr > p:cSldViewPr > p:cViewPr
        $objWriter->startElement('p:cViewPr');

        // p:viewPr > p:slideViewPr > p:cSldViewPr > p:cViewPr > p:scale
        $objWriter->startElement('p:scale');

        $objWriter->startElement('a:sx');
        $objWriter->writeAttribute('d', '100');
        $objWriter->writeAttribute('n', (int)($this->getPresentation()->getPresentationProperties()->getZoom() * 100));
        $objWriter->endElement();

        $objWriter->startElement('a:sy');
        $objWriter->writeAttribute('d', '100');
        $objWriter->writeAttribute('n', (int)($this->getPresentation()->getPresentationProperties()->getZoom() * 100));
        $objWriter->endElement();

        // > // p:viewPr > p:slideViewPr > p:cSldViewPr > p:cViewPr > p:scale
        $objWriter->endElement();

        $objWriter->startElement('p:origin');
        $objWriter->writeAttribute('x', '0');
        $objWriter->writeAttribute('y', '0');
        $objWriter->endElement();

        // > // p:viewPr > p:slideViewPr > p:cSldViewPr > p:cViewPr
        $objWriter->endElement();

        // > // p:viewPr > p:slideViewPr > p:cSldViewPr
        $objWriter->endElement();

        // > // p:viewPr > p:slideViewPr
        $objWriter->endElement();

        // > // p:viewPr
        $objWriter->endElement();

        $this->getZip()->addFromString('ppt/viewProps.xml', $objWriter->getData());

        return $this->getZip();
    }
}
