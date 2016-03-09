<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\PackDefault;
use PhpOffice\Common\XMLWriter;

class PptTheme extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     * @throws \Exception
     */
    public function render()
    {
        $oLayoutPack = new PackDefault();

        $masterSlides = $oLayoutPack->getMasterSlides();
        foreach ($masterSlides as $masterSlide) {
            // Add themes to ZIP file
            $this->getZip()->addFromString('ppt/theme/_rels/theme' . $masterSlide['masterid'] . '.xml.rels', $this->writeThemeRelationships($masterSlide['masterid']));
            $this->getZip()->addFromString('ppt/theme/theme' . $masterSlide['masterid'] . '.xml', $this->writeTheme($masterSlide['masterid']));
        }

        $otherRelations = $oLayoutPack->getThemeRelations();
        foreach ($otherRelations as $otherRelation) {
            if (strpos($otherRelation['target'], 'http://') !== 0) {
                $this->getZip()->addFromString($this->absoluteZipPath('ppt/theme/' . $otherRelation['target']), $otherRelation['contents']);
            }
        }

        return $this->getZip();
    }


    /**
     * Write theme to XML format
     *
     * @param  int $masterId
     * @throws \Exception
     * @return string XML Output
     */
    public function writeTheme($masterId = 1)
    {
        $oLayoutPack = new PackDefault();
        foreach ($oLayoutPack->getThemes() as $theme) {
            if ($theme['masterid'] == $masterId) {
                return $theme['body'];
            }
        }
        throw new \Exception('No theme has been found!');
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

        // Layout pack
        $oLayoutPack = new PackDefault();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Other relationships
        $otherRelations = $oLayoutPack->getThemeRelations();
        foreach ($otherRelations as $otherRelation) {
            if ($otherRelation['masterid'] == $masterId) {
                $this->writeRelationship($objWriter, $otherRelation['id'], $otherRelation['type'], $otherRelation['target']);
            }
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
