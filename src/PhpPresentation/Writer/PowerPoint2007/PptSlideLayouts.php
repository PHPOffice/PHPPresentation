<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Style\ColorMap;

class PptSlideLayouts extends AbstractSlide
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        foreach ($this->oPresentation->getAllMasterSlides() as $oSlideMaster) {
            foreach ($oSlideMaster->getAllSlideLayouts() as $oSlideLayout) {
                $this->oZip->addFromString('ppt/slideLayouts/_rels/slideLayout' . $oSlideLayout->layoutNr . '.xml.rels', $this->writeSlideLayoutRelationships($oSlideMaster->getRelsIndex()));
                $this->oZip->addFromString('ppt/slideLayouts/slideLayout' . $oSlideLayout->layoutNr . '.xml', $this->writeSlideLayout($oSlideLayout));
            }
        }

        return $this->oZip;
    }


    /**
     * Write slide layout relationships to XML format
     *
     * @param  int       $masterId
     * @return string    XML Output
     * @throws \Exception
     */
    public function writeSlideLayoutRelationships($masterId = 1)
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Write slideMaster relationship
        $this->writeRelationship($objWriter, 1, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', '../slideMasters/slideMaster' . $masterId . '.xml');

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write slide to XML format
     *
     * @param  \PhpOffice\PhpPresentation\Slide\SlideLayout $pSlideLayout
     * @return string XML Output
     * @throws \Exception
     */
    public function writeSlideLayout(SlideLayout $pSlideLayout)
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');
        // p:sldLayout
        $objWriter->startElement('p:sldLayout');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $objWriter->writeAttribute('preserve', 1);
        // p:sldLayout\p:cSld
        $objWriter->startElement('p:cSld');
        $objWriter->writeAttributeIf($pSlideLayout->getLayoutName() != '', 'name', $pSlideLayout->getLayoutName());
        // Background
        $this->writeSlideBackground($pSlideLayout, $objWriter);
        // p:sldLayout\p:cSld\p:spTree
        $objWriter->startElement('p:spTree');
        // p:sldLayout\p:cSld\p:spTree\p:nvGrpSpPr
        $objWriter->startElement('p:nvGrpSpPr');
        // p:sldLayout\p:cSld\p:spTree\p:nvGrpSpPr\p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '1');
        $objWriter->writeAttribute('name', '');
        $objWriter->endElement();
        // p:sldLayout\p:cSld\p:spTree\p:nvGrpSpPr\p:cNvGrpSpPr
        $objWriter->writeElement('p:cNvGrpSpPr', null);
        // p:sldLayout\p:cSld\p:spTree\p:nvGrpSpPr\p:nvPr
        $objWriter->writeElement('p:nvPr', null);
        // p:sldLayout\p:cSld\p:spTree\p:nvGrpSpPr
        $objWriter->endElement();
        // p:sldLayout\p:cSld\p:spTree\p:grpSpPr
        $objWriter->startElement('p:grpSpPr');
        // p:sldLayout\p:cSld\p:spTree\p:grpSpPr\a:xfrm
        $objWriter->startElement('a:xfrm');
        // p:sldLayout\p:cSld\p:spTree\p:grpSpPr\a:xfrm\a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pSlideLayout->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pSlideLayout->getOffsetY()));
        $objWriter->endElement();
        // p:sldLayout\p:cSld\p:spTree\p:grpSpPr\a:xfrm\a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pSlideLayout->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pSlideLayout->getExtentY()));
        $objWriter->endElement();
        // p:sldLayout\p:cSld\p:spTree\p:grpSpPr\a:xfrm\a:chOff
        $objWriter->startElement('a:chOff');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pSlideLayout->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pSlideLayout->getOffsetY()));
        $objWriter->endElement();
        // p:sldLayout\p:cSld\p:spTree\p:grpSpPr\a:xfrm\a:chExt
        $objWriter->startElement('a:chExt');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pSlideLayout->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pSlideLayout->getExtentY()));
        $objWriter->endElement();
        // p:sldLayout\p:cSld\p:spTree\p:grpSpPr\a:xfrm\
        $objWriter->endElement();
        // p:sldLayout\p:cSld\p:spTree\p:grpSpPr\
        $objWriter->endElement();

        // Loop shapes
        $this->writeShapeCollection($objWriter, $pSlideLayout->getShapeCollection());
        // p:sldLayout\p:cSld\p:spTree\
        $objWriter->endElement();
        // p:sldLayout\p:cSld\
        $objWriter->endElement();

        // p:sldLayout\p:clrMapOvr
        $objWriter->startElement('p:clrMapOvr');
        $arrayDiff = array_diff_assoc(ColorMap::$mappingDefault, $pSlideLayout->colorMap->getMapping());
        if (!empty($arrayDiff)) {
            // p:sldLayout\p:clrMapOvr\a:overrideClrMapping
            $objWriter->startElement('a:overrideClrMapping');
            foreach ($pSlideLayout->colorMap->getMapping() as $n => $v) {
                $objWriter->writeAttribute($n, $v);
            }
            $objWriter->endElement();
        } else {
            // p:sldLayout\p:clrMapOvr\a:masterClrMapping
            $objWriter->writeElement('a:masterClrMapping');
        }
        // p:sldLayout\p:clrMapOvr\
        $objWriter->endElement();

        if (!is_null($pSlideLayout->getTransition())) {
            $this->writeSlideTransition($objWriter, $pSlideLayout->getTransition());
        }

        // p:sldLayout\
        $objWriter->endElement();

        return $objWriter->getData();
    }
}
