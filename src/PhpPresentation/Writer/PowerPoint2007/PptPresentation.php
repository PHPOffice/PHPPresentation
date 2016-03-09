<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\PackDefault;

class PptPresentation extends AbstractDecoratorWriter
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

        // p:presentation
        $objWriter->startElement('p:presentation');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // p:sldMasterIdLst
        $objWriter->startElement('p:sldMasterIdLst');

        // Add slide masters
        $relationId    = 1;
        $slideMasterId = 2147483648;
        $oLayoutPack = new PackDefault();
        $masterSlides  = $oLayoutPack->getMasterSlides();
        $masterSlidesCount = count($masterSlides);
        // @todo foreach ($masterSlides as $masterSlide)
        for ($i = 0; $i < $masterSlidesCount; $i++) {
            // p:sldMasterId
            $objWriter->startElement('p:sldMasterId');
            $objWriter->writeAttribute('id', $slideMasterId);
            $objWriter->writeAttribute('r:id', 'rId' . $relationId++);
            $objWriter->endElement();

            // Increase identifier
            $slideMasterId += 12;
        }
        $objWriter->endElement();

        // theme
        $relationId++;

        // p:sldIdLst
        $objWriter->startElement('p:sldIdLst');
        // Write slides
        $slideCount = $this->oPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            // p:sldId
            $objWriter->startElement('p:sldId');
            $objWriter->writeAttribute('id', ($i + 256));
            $objWriter->writeAttribute('r:id', 'rId' . ($i + $relationId));
            $objWriter->endElement();
        }
        $objWriter->endElement();

        // p:sldSz
        $objWriter->startElement('p:sldSz');
        $objWriter->writeAttribute('cx', $this->oPresentation->getLayout()->getCX());
        $objWriter->writeAttribute('cy', $this->oPresentation->getLayout()->getCY());
        if ($this->oPresentation->getLayout()->getDocumentLayout() != DocumentLayout::LAYOUT_CUSTOM) {
            $objWriter->writeAttribute('type', $this->oPresentation->getLayout()->getDocumentLayout());
        }
        $objWriter->endElement();

        // p:notesSz
        $objWriter->startElement('p:notesSz');
        $objWriter->writeAttribute('cx', '6858000');
        $objWriter->writeAttribute('cy', '9144000');
        $objWriter->endElement();

        $objWriter->endElement();

        $this->oZip->addFromString('ppt/presentation.xml', $objWriter->getData());

        // Return
        return $this->oZip;
    }
}
