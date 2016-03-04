<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\PackDefault;

class PptSlideLayouts extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        $oLayoutPack = new PackDefault();

        // Add slide layouts to ZIP file
        foreach ($oLayoutPack->getLayouts() as $key => $layout) {
            $this->oZip->addFromString('ppt/slideLayouts/_rels/slideLayout' . $key . '.xml.rels', $this->writeSlideLayoutRelationships($key, $layout['masterid']));
            $this->oZip->addFromString('ppt/slideLayouts/slideLayout' . $key . '.xml', utf8_encode($layout['body']));
        }

        foreach ($oLayoutPack->getLayoutRelations() as $otherRelation) {
            if (strpos($otherRelation['target'], 'http://') !== 0) {
                $this->oZip->addFromString($this->absoluteZipPath('ppt/slideLayouts/' . $otherRelation['target']), $otherRelation['contents']);
            }
        }

        return $this->oZip;
    }


    /**
     * Write slide layout relationships to XML format
     *
     * @param  int       $slideLayoutIndex
     * @param  int       $masterId
     * @return string    XML Output
     * @throws \Exception
     */
    public function writeSlideLayoutRelationships($slideLayoutIndex, $masterId = 1)
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // Layout pack
        $oLayoutPack = new PackDefault();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Write slideMaster relationship
        $this->writeRelationship($objWriter, 1, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', '../slideMasters/slideMaster' . $masterId . '.xml');

        // Other relationships
        foreach ($oLayoutPack->getLayoutRelations() as $otherRelation) {
            if ($otherRelation['layoutId'] == $slideLayoutIndex) {
                $this->writeRelationship($objWriter, $otherRelation['id'], $otherRelation['type'], $otherRelation['target']);
            }
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
