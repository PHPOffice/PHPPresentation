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
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Style\ColorMap;

class PptSlideLayouts extends AbstractSlide
{
    public function render(): ZipInterface
    {
        foreach ($this->oPresentation->getAllMasterSlides() as $oSlideMaster) {
            foreach ($oSlideMaster->getAllSlideLayouts() as $oSlideLayout) {
                $this->oZip->addFromString('ppt/slideLayouts/_rels/slideLayout' . $oSlideLayout->layoutNr . '.xml.rels', $this->writeSlideLayoutRelationships($oSlideLayout));
                $this->oZip->addFromString('ppt/slideLayouts/slideLayout' . $oSlideLayout->layoutNr . '.xml', $this->writeSlideLayout($oSlideLayout));

                // Add background image slide
                $oBkgImage = $oSlideLayout->getBackground();
                if ($oBkgImage instanceof Image) {
                    $this->oZip->addFromString('ppt/media/' . $oBkgImage->getIndexedFilename($oSlideLayout->getRelsIndex()), file_get_contents($oBkgImage->getPath()));
                }
            }
        }

        return $this->oZip;
    }

    /**
     * Write slide layout relationships to XML format.
     *
     * @return string XML Output
     */
    protected function writeSlideLayoutRelationships(SlideLayout $oSlideLayout): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        $relId = 0;

        // Write slideMaster relationship
        $this->writeRelationship($objWriter, ++$relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster', '../slideMasters/slideMaster' . $oSlideLayout->getSlideMaster()->getRelsIndex() . '.xml');

        // Write drawing relationships?
        $relId = $this->writeDrawingRelations($oSlideLayout, $objWriter, ++$relId);

        // Write background relationships?
        $oBackground = $oSlideLayout->getBackground();
        if ($oBackground instanceof Image) {
            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $oBackground->getIndexedFilename($oSlideLayout->getRelsIndex()));
            $oBackground->relationId = 'rId' . $relId;

            ++$relId;
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write slide to XML format.
     *
     * @return string XML Output
     */
    protected function writeSlideLayout(SlideLayout $pSlideLayout): string
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
        $objWriter->writeAttributeIf('' != $pSlideLayout->getLayoutName(), 'name', $pSlideLayout->getLayoutName());
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

        if (null !== $pSlideLayout->getTransition()) {
            $this->writeSlideTransition($objWriter, $pSlideLayout->getTransition());
        }

        // p:sldLayout\
        $objWriter->endElement();

        return $objWriter->getData();
    }
}
