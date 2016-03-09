<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\PackDefault;

class PptSlideMasters extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        $oLayoutPack = new PackDefault();

        foreach ($oLayoutPack->getMasterSlides() as $masterSlide) {
            // Add slide masters to ZIP file
            $this->oZip->addFromString('ppt/slideMasters/_rels/slideMaster' . $masterSlide['masterid'] . '.xml.rels', $this->writeSlideMasterRelationships($masterSlide['masterid']));
            $this->oZip->addFromString('ppt/slideMasters/slideMaster' . $masterSlide['masterid'] . '.xml', $masterSlide['body']);
        }

        // Add layoutpack relations
        $otherRelations = $oLayoutPack->getMasterSlideRelations();
        foreach ($otherRelations as $otherRelation) {
            if (strpos($otherRelation['target'], 'http://') !== 0) {
                $this->oZip->addFromString($this->absoluteZipPath('ppt/slideMasters/' . $otherRelation['target']), $otherRelation['contents']);
            }
        }

        return $this->oZip;
    }

    /**
     * Write slide master relationships to XML format
     *
     * @param  int       $masterId Master slide id
     * @return string    XML Output
     * @throws \Exception
     */
    public function writeSlideMasterRelationships($masterId = 1)
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Keep content id
        $contentId = 0;

        // Lookup layouts
        $layouts    = array();
        $oLayoutPack = new PackDefault();
        foreach ($oLayoutPack->getLayouts() as $key => $layout) {
            if ($layout['masterid'] == $masterId) {
                $layouts[$key] = $layout;
            }
        }

        // Write slideLayout relationships
        foreach ($layouts as $key => $layout) {
            $this->writeRelationship($objWriter, ++$contentId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout' . $key . '.xml');
        }

        // Relationship theme/theme1.xml
        $this->writeRelationship($objWriter, ++$contentId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', '../theme/theme' . $masterId . '.xml');

        // Other relationships
        $otherRelations = $oLayoutPack->getMasterSlideRelations();
        foreach ($otherRelations as $otherRelation) {
            if ($otherRelation['masterid'] == $masterId) {
                $this->writeRelationship($objWriter, ++$contentId, $otherRelation['type'], $otherRelation['target']);
            }
        }
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
