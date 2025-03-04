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
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Style\SchemeColor;

class PptSlideMasters extends AbstractSlide
{
    public function render(): ZipInterface
    {
        foreach ($this->oPresentation->getAllMasterSlides() as $oMasterSlide) {
            // Add the relations from the masterSlide to the ZIP file
            $this->oZip->addFromString('ppt/slideMasters/_rels/slideMaster' . $oMasterSlide->getRelsIndex() . '.xml.rels', $this->writeSlideMasterRelationships($oMasterSlide));
            // Add the information from the masterSlide to the ZIP file
            $this->oZip->addFromString('ppt/slideMasters/slideMaster' . $oMasterSlide->getRelsIndex() . '.xml', $this->writeSlideMaster($oMasterSlide));

            // Add background image slide
            $oBkgImage = $oMasterSlide->getBackground();
            if ($oBkgImage instanceof Image) {
                $this->oZip->addFromString('ppt/media/' . $oBkgImage->getIndexedFilename($oMasterSlide->getRelsIndex()), file_get_contents($oBkgImage->getPath()));
            }
        }

        return $this->oZip;
    }

    /**
     * Write slide master relationships to XML format.
     *
     * @todo Set method in protected
     *
     * @return string XML Output
     */
    public function writeSlideMasterRelationships(SlideMaster $oMasterSlide): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');
        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        // Starting relation id
        $relId = 0;
        // Write all the relations to the Layout Slides
        foreach ($oMasterSlide->getAllSlideLayouts() as $slideLayout) {
            $this->writeRelationship($objWriter, ++$relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout' . $slideLayout->layoutNr . '.xml');
            // Save the used relationId
            $slideLayout->relationId = 'rId' . $relId;
        }

        // Write drawing relationships?
        $relId = $this->writeDrawingRelations($oMasterSlide, $objWriter, ++$relId);

        // Write background relationships?
        $oBackground = $oMasterSlide->getBackground();
        if ($oBackground instanceof Image) {
            $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', '../media/' . $oBackground->getIndexedFilename($oMasterSlide->getRelsIndex()));
            $oBackground->relationId = 'rId' . $relId;

            ++$relId;
        }

        // TODO: Write hyperlink relationships?
        // TODO: Write comment relationships
        // Relationship theme/theme1.xml
        $this->writeRelationship($objWriter, $relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', '../theme/theme' . $oMasterSlide->getRelsIndex() . '.xml');
        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * Write slide to XML format.
     *
     * @return string XML Output
     */
    protected function writeSlideMaster(SlideMaster $pSlide): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');
        // p:sldMaster
        $objWriter->startElement('p:sldMaster');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        // p:sldMaster\p:cSld
        $objWriter->startElement('p:cSld');
        // Background
        $this->writeSlideBackground($pSlide, $objWriter);
        // p:sldMaster\p:cSld\p:spTree
        $objWriter->startElement('p:spTree');
        // p:sldMaster\p:cSld\p:spTree\p:nvGrpSpPr
        $objWriter->startElement('p:nvGrpSpPr');
        // p:sldMaster\p:cSld\p:spTree\p:nvGrpSpPr\p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '1');
        $objWriter->writeAttribute('name', '');
        $objWriter->endElement();
        // p:sldMaster\p:cSld\p:spTree\p:nvGrpSpPr\p:cNvGrpSpPr
        $objWriter->writeElement('p:cNvGrpSpPr', null);
        // p:sldMaster\p:cSld\p:spTree\p:nvGrpSpPr\p:nvPr
        $objWriter->writeElement('p:nvPr', null);
        // p:sldMaster\p:cSld\p:spTree\p:nvGrpSpPr
        $objWriter->endElement();
        // p:sldMaster\p:cSld\p:spTree\p:grpSpPr
        $objWriter->startElement('p:grpSpPr');
        // p:sldMaster\p:cSld\p:spTree\p:grpSpPr\a:xfrm
        $objWriter->startElement('a:xfrm');
        // p:sldMaster\p:cSld\p:spTree\p:grpSpPr\a:xfrm\a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', 0);
        $objWriter->writeAttribute('y', 0);
        $objWriter->endElement();
        // p:sldMaster\p:cSld\p:spTree\p:grpSpPr\a:xfrm\a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', 0);
        $objWriter->writeAttribute('cy', 0);
        $objWriter->endElement();
        // p:sldMaster\p:cSld\p:spTree\p:grpSpPr\a:xfrm\a:chOff
        $objWriter->startElement('a:chOff');
        $objWriter->writeAttribute('x', 0);
        $objWriter->writeAttribute('y', 0);
        $objWriter->endElement();
        // p:sldMaster\p:cSld\p:spTree\p:grpSpPr\a:xfrm\a:chExt
        $objWriter->startElement('a:chExt');
        $objWriter->writeAttribute('cx', 0);
        $objWriter->writeAttribute('cy', 0);
        $objWriter->endElement();
        // p:sldMaster\p:cSld\p:spTree\p:grpSpPr\a:xfrm\
        $objWriter->endElement();
        // p:sldMaster\p:cSld\p:spTree\p:grpSpPr\
        $objWriter->endElement();
        // Loop shapes
        $this->writeShapeCollection($objWriter, $pSlide->getShapeCollection());
        // p:sldMaster\p:cSld\p:spTree\
        $objWriter->endElement();
        // p:sldMaster\p:cSld\
        $objWriter->endElement();

        // p:sldMaster\p:clrMap
        $objWriter->startElement('p:clrMap');
        foreach ($pSlide->colorMap->getMapping() as $n => $v) {
            $objWriter->writeAttribute($n, $v);
        }
        $objWriter->endElement();
        // p:sldMaster\p:clrMap\

        // p:sldMaster\p:sldLayoutIdLst
        $objWriter->startElement('p:sldLayoutIdLst');
        foreach ($pSlide->getAllSlideLayouts() as $layout) {
            // @var $layout Slide\SlideLayout
            $objWriter->startElement('p:sldLayoutId');
            $objWriter->writeAttribute('id', $layout->layoutId);
            $objWriter->writeAttribute('r:id', $layout->relationId);
            $objWriter->endElement();
        }
        $objWriter->endElement();
        // p:sldMaster\p:sldLayoutIdLst\

        // p:sldMaster\p:txStyles
        $objWriter->startElement('p:txStyles');
        foreach ([
            'p:titleStyle' => $pSlide->getTextStyles()->getTitleStyle(),
            'p:bodyStyle' => $pSlide->getTextStyles()->getBodyStyle(),
            'p:otherStyle' => $pSlide->getTextStyles()->getOtherStyle(),
        ] as $startElement => $stylesArray) {
            // titleStyle
            $objWriter->startElement($startElement);
            foreach ($stylesArray as $lvl => $oParagraph) {
                /** @var RichText\Paragraph $oParagraph */
                $elementName = (0 == $lvl ? 'a:defPPr' : 'a:lvl' . $lvl . 'pPr');
                $objWriter->startElement($elementName);
                $objWriter->writeAttribute('algn', $oParagraph->getAlignment()->getHorizontal());
                $objWriter->writeAttributeIf(
                    0 != $oParagraph->getAlignment()->getMarginLeft(),
                    'marL',
                    CommonDrawing::pixelsToEmu($oParagraph->getAlignment()->getMarginLeft())
                );
                $objWriter->writeAttributeIf(
                    0 != $oParagraph->getAlignment()->getMarginRight(),
                    'marR',
                    CommonDrawing::pixelsToEmu($oParagraph->getAlignment()->getMarginRight())
                );
                $objWriter->writeAttributeIf(
                    0 != $oParagraph->getAlignment()->getIndent(),
                    'indent',
                    (int) CommonDrawing::pixelsToEmu($oParagraph->getAlignment()->getIndent())
                );
                $objWriter->startElement('a:defRPr');
                $objWriter->writeAttributeIf(10 != $oParagraph->getFont()->getSize(), 'sz', $oParagraph->getFont()->getSize() * 100);
                $objWriter->writeAttributeIf($oParagraph->getFont()->isBold(), 'b', 1);
                $objWriter->writeAttributeIf($oParagraph->getFont()->isItalic(), 'i', 1);
                $objWriter->writeAttribute('kern', '1200');
                if ($oParagraph->getFont()->getColor() instanceof SchemeColor) {
                    $objWriter->startElement('a:solidFill');
                    $objWriter->startElement('a:schemeClr');
                    $objWriter->writeAttribute('val', $oParagraph->getFont()->getColor()->getValue());
                    $objWriter->endElement();
                    $objWriter->endElement();
                }
                $objWriter->endElement();
                $objWriter->endElement();
            }
            $objWriter->writeElement('a:extLst');
            $objWriter->endElement();
        }
        $objWriter->endElement();
        // p:sldMaster\p:txStyles\

        if (null !== $pSlide->getTransition()) {
            $this->writeSlideTransition($objWriter, $pSlide->getTransition());
        }

        // p:sldMaster\
        $objWriter->endElement();

        return $objWriter->getData();
    }
}
