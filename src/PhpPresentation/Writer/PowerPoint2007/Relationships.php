<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Comment\Author;

class Relationships extends AbstractDecoratorWriter
{
    /**
     * Add relationships to ZIP file.
     */
    public function render(): ZipInterface
    {
        $this->getZip()->addFromString('_rels/.rels', $this->writeRelationships());
        $this->getZip()->addFromString('ppt/_rels/presentation.xml.rels', $this->writePresentationRelationships());

        return $this->getZip();
    }

    /**
     * Write relationships to XML format.
     *
     * @return string XML Output
     */
    protected function writeRelationships(): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        // Relationship ppt/presentation.xml
        $this->writeRelationship($objWriter, 1, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'ppt/presentation.xml');
        // Relationship docProps/core.xml
        $this->writeRelationship($objWriter, 2, 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml');
        // Relationship docProps/app.xml
        $this->writeRelationship($objWriter, 3, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'docProps/app.xml');
        // Relationship docProps/custom.xml
        $this->writeRelationship($objWriter, 4, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties', 'docProps/custom.xml');

        $idxRelation = 5;

        // Relationship docProps/thumbnail.jpeg
        $thumnbail = $this->getPresentation()->getPresentationProperties()->getThumbnail();
        if ($thumnbail) {
            $gdImage = imagecreatefromstring($thumnbail);
            if ($gdImage) {
                imagedestroy($gdImage);
                // Relationship docProps/thumbnail.jpeg
                $this->writeRelationship($objWriter, $idxRelation, 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail', 'docProps/thumbnail.jpeg');
                // $idxRelation++;
            }
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write presentation relationships to XML format.
     *
     * @return string XML Output
     */
    protected function writePresentationRelationships(): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        // Relation id
        $relationId = 1;

        foreach ($this->getPresentation()->getAllMasterSlides() as $oMasterSlide) {
            // Relationship slideMasters/slideMasterX.xml
            $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', 'slideMasters/slideMaster' . $oMasterSlide->getRelsIndex() . '.xml');
        }

        // Add slide theme (only one!)
        // Relationship theme/theme1.xml
        $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', 'theme/theme1.xml');

        // Relationships with slides
        $slideCount = $this->getPresentation()->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', 'slides/slide' . ($i + 1) . '.xml');
        }

        $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps', 'presProps.xml');
        $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps', 'viewProps.xml');
        $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles', 'tableStyles.xml');

        // Comments Authors
        foreach ($this->getPresentation()->getAllSlides() as $oSlide) {
            foreach ($oSlide->getShapeCollection() as $oShape) {
                if (!($oShape instanceof Comment)) {
                    continue;
                }
                $oAuthor = $oShape->getAuthor();
                if (!($oAuthor instanceof Author)) {
                    continue;
                }
                $this->writeRelationship($objWriter, $relationId++, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors', 'commentAuthors.xml');

                break 2;
            }
        }
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
