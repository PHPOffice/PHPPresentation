<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\DocumentLayout;

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

        $countMasterSlides = count($this->oPresentation->getAllMasterSlides());
        for ($inc = 1; $inc <= $countMasterSlides; $inc++) {
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

        $objWriter->writeRaw('<p:defaultTextStyle>
  <a:defPPr>
   <a:defRPr lang="fr-FR"/>
  </a:defPPr>
  <a:lvl1pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="0" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl1pPr>
  <a:lvl2pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="457200" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl2pPr>
  <a:lvl3pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="914400" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl3pPr>
  <a:lvl4pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="1371600" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl4pPr>
  <a:lvl5pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="1828800" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl5pPr>
  <a:lvl6pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="2286000" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl6pPr>
  <a:lvl7pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="2743200" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl7pPr>
  <a:lvl8pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="3200400" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl8pPr>
  <a:lvl9pPr algn="l" defTabSz="914400" eaLnBrk="1" hangingPunct="1" latinLnBrk="0" marL="3657600" rtl="0">
   <a:defRPr kern="1200" sz="1800">
    <a:solidFill>
     <a:schemeClr val="tx1"/>
    </a:solidFill>
    <a:latin typeface="+mn-lt"/>
    <a:ea typeface="+mn-ea"/>
    <a:cs typeface="+mn-cs"/>
   </a:defRPr>
  </a:lvl9pPr>
 </p:defaultTextStyle>
');

        $objWriter->endElement();

        $this->oZip->addFromString('ppt/presentation.xml', $objWriter->getData());

        // Return
        return $this->oZip;
    }
}
