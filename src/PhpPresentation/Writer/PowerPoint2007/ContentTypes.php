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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Chart as ShapeChart;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Drawing as ShapeDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007\ContentTypes
 */
class ContentTypes extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     * @throws \Exception
     */
    public function render()
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Types
        $objWriter->startElement('Types');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // Rels
        $this->writeDefaultContentType($objWriter, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');

        // XML
        $this->writeDefaultContentType($objWriter, 'xml', 'application/xml');

        // Presentation
        $this->writeOverrideContentType($objWriter, '/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml');

        // PptProps
        $this->writeOverrideContentType($objWriter, '/ppt/presProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml');
        $this->writeOverrideContentType($objWriter, '/ppt/tableStyles.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml');
        $this->writeOverrideContentType($objWriter, '/ppt/viewProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml');

        // DocProps
        $this->writeOverrideContentType($objWriter, '/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
        $this->writeOverrideContentType($objWriter, '/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');
        $this->writeOverrideContentType($objWriter, '/docProps/custom.xml', 'application/vnd.openxmlformats-officedocument.custom-properties+xml');

        // Slide masters
        $sldLayoutNr = 0;
        $sldLayoutId = time() + 689016272; // requires minimum value of 2 147 483 648
        foreach ($this->oPresentation->getAllMasterSlides() as $idx => $oSlideMaster) {
            $oSlideMaster->setRelsIndex($idx + 1);
            $this->writeOverrideContentType($objWriter, '/ppt/slideMasters/slideMaster' . $oSlideMaster->getRelsIndex() . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml');
            $this->writeOverrideContentType($objWriter, '/ppt/theme/theme' . $oSlideMaster->getRelsIndex() . '.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
            foreach ($oSlideMaster->getAllSlideLayouts() as $oSlideLayout) {
                $oSlideLayout->layoutNr = ++$sldLayoutNr;
                $oSlideLayout->layoutId = ++$sldLayoutId;
                $this->writeOverrideContentType($objWriter, '/ppt/slideLayouts/slideLayout' . $oSlideLayout->layoutNr . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml');
            }
        }

        // Slides
        $hasComments = false;
        $slideCount = $this->oPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            $oSlide = $this->oPresentation->getSlide($i);
            $this->writeOverrideContentType($objWriter, '/ppt/slides/slide' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
            if ($oSlide->getNote()->getShapeCollection()->count() > 0) {
                $this->writeOverrideContentType($objWriter, '/ppt/notesSlides/notesSlide' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
            }
            foreach ($oSlide->getShapeCollection() as $oShape) {
                if ($oShape instanceof Comment) {
                    $this->writeOverrideContentType($objWriter, '/ppt/comments/comment' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');
                    $hasComments = true;
                    break;
                }
            }
        }

        if ($hasComments) {
            $this->writeOverrideContentType($objWriter, '/ppt/commentAuthors.xml', 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml');
        }

        // Add media content-types
        $aMediaContentTypes = array();

        // GIF, JPEG, PNG
        $aMediaContentTypes['gif']  = 'image/gif';
        $aMediaContentTypes['jpg']  = 'image/jpeg';
        $aMediaContentTypes['jpeg'] = 'image/jpeg';
        $aMediaContentTypes['png']  = 'image/png';
        foreach ($aMediaContentTypes as $key => $value) {
            $this->writeDefaultContentType($objWriter, $key, $value);
        }

        // XLSX
        $this->writeDefaultContentType($objWriter, 'xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        // Other media content types
        $mediaCount = $this->getDrawingHashTable()->count();
        for ($i = 0; $i < $mediaCount; ++$i) {
            $shapeIndex = $this->getDrawingHashTable()->getByIndex($i);
            if ($shapeIndex instanceof ShapeChart) {
                // Chart content type
                $this->writeOverrideContentType($objWriter, '/ppt/charts/chart' . $shapeIndex->getImageIndex() . '.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
            } else {
                $extension = strtolower($shapeIndex->getExtension());
                $mimeType = $shapeIndex->getMimeType();

                if (!isset($aMediaContentTypes[$extension])) {
                    $aMediaContentTypes[$extension] = $mimeType;

                    $this->writeDefaultContentType($objWriter, $extension, $mimeType);
                }
            }
        }

        $objWriter->endElement();

        $this->oZip->addFromString('[Content_Types].xml', $objWriter->getData());

        return $this->oZip;
    }

    /**
     * Write Default content type
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter    XML Writer
     * @param  string                         $pPartname    Part name
     * @param  string                         $pContentType Content type
     * @throws \Exception
     */
    private function writeDefaultContentType(XMLWriter $objWriter, $pPartname = '', $pContentType = '')
    {
        if ($pPartname == '' || $pContentType == '') {
            throw new \Exception("Invalid parameters passed.");
        }
        // Write content type
        $objWriter->startElement('Default');
        $objWriter->writeAttribute('Extension', $pPartname);
        $objWriter->writeAttribute('ContentType', $pContentType);
        $objWriter->endElement();
    }

    /**
     * Write Override content type
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter    XML Writer
     * @param  string                         $pPartname    Part name
     * @param  string                         $pContentType Content type
     * @throws \Exception
     */
    private function writeOverrideContentType(XMLWriter $objWriter, $pPartname = '', $pContentType = '')
    {
        if ($pPartname == '' || $pContentType == '') {
            throw new \Exception("Invalid parameters passed.");
        }
        // Write content type
        $objWriter->startElement('Override');
        $objWriter->writeAttribute('PartName', $pPartname);
        $objWriter->writeAttribute('ContentType', $pContentType);
        $objWriter->endElement();
    }
}
