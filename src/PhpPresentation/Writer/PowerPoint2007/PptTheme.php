<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Slide;
use PhpOffice\Common\XMLWriter;

class PptTheme extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     * @throws \Exception
     */
    public function render()
    {
        foreach ($this->oPresentation->getAllMasterSlides() as $oMasterSlide) {
            $this->getZip()->addFromString('ppt/theme/_rels/theme' . $oMasterSlide->getRelsIndex() . '.xml.rels', $this->writeThemeRelationships($oMasterSlide->getRelsIndex()));
            $this->getZip()->addFromString('ppt/theme/theme' . $oMasterSlide->getRelsIndex() . '.xml', $this->writeTheme($oMasterSlide));
        }

        /*
        $otherRelations = $oLayoutPack->getThemeRelations();
        foreach ($otherRelations as $otherRelation) {
            if (strpos($otherRelation['target'], 'http://') !== 0) {
                $this->getZip()->addFromString($this->absoluteZipPath('ppt/theme/' . $otherRelation['target']), $otherRelation['contents']);
            }
        }
        */

        return $this->getZip();
    }


    /**
     * Write theme to XML format
     *
     * @param  Slide\SlideMaster $oMasterSlide
     * @return string XML Output
     */
    protected function writeTheme(Slide\SlideMaster $oMasterSlide)
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        $name = 'Theme'.rand(1, 100);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // a:theme
        $objWriter->startElement('a:theme');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('name', $name);

        // a:theme/a:themeElements
        $objWriter->startElement('a:themeElements');

        // a:theme/a:themeElements/a:clrScheme
        $objWriter->startElement('a:clrScheme');
        $objWriter->writeAttribute('name', $name);

        foreach ($oMasterSlide->getAllSchemeColors() as $oSchemeColor) {
            // a:theme/a:themeElements/a:clrScheme/a:*
            $objWriter->startElement('a:'.$oSchemeColor->getValue());

            if (in_array($oSchemeColor->getValue(), array(
                'dk1', 'lt1'
            ))) {
                $objWriter->startElement('a:sysClr');
                $objWriter->writeAttribute('val', ($oSchemeColor->getValue() == 'dk1' ? 'windowText' : 'window'));
                $objWriter->writeAttribute('lastClr', $oSchemeColor->getRGB());
                $objWriter->endElement();
            } else {
                $objWriter->startElement('a:srgbClr');
                $objWriter->writeAttribute('val', $oSchemeColor->getRGB());
                $objWriter->endElement();
            }

            // a:theme/a:themeElements/a:clrScheme/a:*/
            $objWriter->endElement();
        }

        // a:theme/a:themeElements/a:clrScheme/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme
        $objWriter->startElement('a:fmtScheme');
        $objWriter->writeAttribute('name', $name);

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst
        $objWriter->startElement('a:bgFillStyleLst');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill
        $objWriter->startElement('a:solidFill');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/
        $objWriter->endElement();

        // a:theme/a:themeElements/
        $objWriter->endElement();

        // a:theme/a:themeElements
        $objWriter->writeElement('a:objectDefaults');

        // a:theme/a:extraClrSchemeLst
        $objWriter->writeElement('a:extraClrSchemeLst');

        // a:theme/
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }


    /**
     * Write theme relationships to XML format
     *
     * @param  int       $masterId Master slide id
     * @return string    XML Output
     * @throws \Exception
     */
    public function writeThemeRelationships($masterId = 1)
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Other relationships
        /*
        $otherRelations = $oLayoutPack->getThemeRelations();
        foreach ($otherRelations as $otherRelation) {
            if ($otherRelation['masterid'] == $masterId) {
                $this->writeRelationship($objWriter, $otherRelation['id'], $otherRelation['type'], $otherRelation['target']);
            }
        }
        */
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
