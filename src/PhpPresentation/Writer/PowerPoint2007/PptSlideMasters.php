<?php
namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\AbstractDrawing;
use PhpOffice\PhpPresentation\Shape\Chart as ShapeChart;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table as ShapeTable;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Style\TextStyle;

class PptSlideMasters extends AbstractSlide
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        $prevLayouts = 0;
        foreach ($this->oPresentation->getAllMasterSlides() as $idx => $oSlide) {
            $oSlide->setRelsIndex($idx + 1);
            // Add the relations from the masterSlide to the ZIP file
            $this->oZip->addFromString('ppt/slideMasters/_rels/slideMaster' . ($idx + 1) . '.xml.rels', $this->writeSlideMasterRelationships($oSlide, $prevLayouts));
            // Add the information from the masterSlide to the ZIP file
            $this->oZip->addFromString('ppt/slideMasters/slideMaster' . ($idx + 1) . '.xml', $this->writeSlideMaster($oSlide));
            $prevLayouts += count($oSlide->getAllSlideLayouts());
        }
//        // Add layoutpack relations
//        $otherRelations = $oLayoutPack->getMasterSlideRelations();
//        foreach ($otherRelations as $otherRelation) {
//            if (strpos($otherRelation['target'], 'http://') !== 0) {
//                $this->oZip->addFromString(
//                  $this->absoluteZipPath('ppt/slideMasters/' . $otherRelation['target']),
//                  $otherRelation['contents']
//              );
//            }
//        }
        return $this->oZip;
    }

    /**
     * Write slide master relationships to XML format
     *
     * @param SlideMaster $pSlideMaster
     * @param int $layoutNr
     * @return string XML Output
     * @throws \Exception
     * @internal param int $masterId Master slide id
     */
    public function writeSlideMasterRelationships(SlideMaster $pSlideMaster, $layoutNr)
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
        foreach ($pSlideMaster->getAllSlideLayouts() as $slideLayout) {
            $this->writeRelationship($objWriter, ++$relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout', '../slideLayouts/slideLayout' . ++$layoutNr . '.xml');
            // Save the used relationId & layout number
            $slideLayout->relationId = 'rId' . $relId;
            $slideLayout->layoutNr = $layoutNr;
        }
        // Write drawing relationships?
        $this->writeDrawingRelations($pSlideMaster, $objWriter, $relId);
        // TODO: Write hyperlink relationships?
        // TODO: Write comment relationships
        // Relationship theme/theme1.xml
        $this->writeRelationship($objWriter, ++$relId, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', '../theme/theme' . $pSlideMaster->getRelsIndex() . '.xml');
        $objWriter->endElement();
        // Return
        return $objWriter->getData();
    }

    /**
     * Write slide to XML format
     *
     * @param  \PhpOffice\PhpPresentation\Slide\SlideMaster $pSlideMaster
     * @return string              XML Output
     * @throws \Exception
     */
    public function writeSlideMaster(SlideMaster $pSlide)
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
        // p:cSld
        $objWriter->startElement('p:cSld');
        // Background
        if ($pSlide->getBackground() instanceof Slide\AbstractBackground) {
            $this->writeSlideBackground($pSlide, $objWriter);
        }
        // p:spTree
        $objWriter->startElement('p:spTree');
        // p:nvGrpSpPr
        $objWriter->startElement('p:nvGrpSpPr');
        // p:cNvPr
        $objWriter->startElement('p:cNvPr');
        $objWriter->writeAttribute('id', '1');
        $objWriter->writeAttribute('name', '');
        $objWriter->endElement();
        // p:cNvGrpSpPr
        $objWriter->writeElement('p:cNvGrpSpPr', null);
        // p:nvPr
        $objWriter->writeElement('p:nvPr', null);
        $objWriter->endElement();
        // p:grpSpPr
        $objWriter->startElement('p:grpSpPr');
        // a:xfrm
        $objWriter->startElement('a:xfrm');
        // a:off
        $objWriter->startElement('a:off');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pSlide->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pSlide->getOffsetY()));
        $objWriter->endElement(); // a:off
        // a:ext
        $objWriter->startElement('a:ext');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pSlide->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pSlide->getExtentY()));
        $objWriter->endElement(); // a:ext
        // a:chOff
        $objWriter->startElement('a:chOff');
        $objWriter->writeAttribute('x', CommonDrawing::pixelsToEmu($pSlide->getOffsetX()));
        $objWriter->writeAttribute('y', CommonDrawing::pixelsToEmu($pSlide->getOffsetY()));
        $objWriter->endElement(); // a:chOff
        // a:chExt
        $objWriter->startElement('a:chExt');
        $objWriter->writeAttribute('cx', CommonDrawing::pixelsToEmu($pSlide->getExtentX()));
        $objWriter->writeAttribute('cy', CommonDrawing::pixelsToEmu($pSlide->getExtentY()));
        $objWriter->endElement(); // a:chExt
        $objWriter->endElement();
        $objWriter->endElement();
        // Loop shapes
        $shapeId = 0;
        $shapes = $pSlide->getShapeCollection();
        foreach ($shapes as $shape) {
            // Increment $shapeId
            ++$shapeId;
            // Check type
            if ($shape instanceof RichText) {
                $this->writeShapeText($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof ShapeTable) {
                $this->writeShapeTable($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Line) {
                $this->writeShapeLine($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof ShapeChart) {
                $this->writeShapeChart($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof AbstractDrawing) {
                $this->writeShapePic($objWriter, $shape, $shapeId);
            } elseif ($shape instanceof Group) {
                $this->writeShapeGroup($objWriter, $shape, $shapeId);
            }
        }
        // TODO
        $objWriter->endElement();
        $objWriter->endElement();
        // < p:clrMap
        $objWriter->startElement('p:clrMap');
        foreach ($pSlide->colorMap->getMapping() as $n => $v) {
            $objWriter->writeAttribute($n, $v);
        }
        $objWriter->endElement();
        // < p:clrMap
        // < p:sldLayoutIdLst
        $objWriter->startElement('p:sldLayoutIdLst');
        $sldLayoutId = time() + 689016272; // requires minimum value of 2147483648
        foreach ($pSlide->getAllSlideLayouts() as $layout) {
            /* @var $layout Slide\SlideLayout */
            $objWriter->startElement('p:sldLayoutId');
            $objWriter->writeAttribute('id', $sldLayoutId++);
            $objWriter->writeAttribute('r:id', $layout->relationId);
            $objWriter->endElement();
        }
        $objWriter->endElement();
        // > p:sldLayoutIdLst
        // p:txStyles
        $objWriter->startElement('p:txStyles');
        foreach (array("p:titleStyle" => $pSlide->getTextStyles()->getTitleStyle(),
                     "p:bodyStyle" => $pSlide->getTextStyles()->getBodyStyle(),
                     "p:otherStyle" => $pSlide->getTextStyles()->getOtherStyle()) as $startElement => $stylesArray) {
            // titleStyle
            $objWriter->startElement($startElement);
            foreach ($stylesArray as $lvl => $oParagraph) {
                /** @var RichText\Paragraph $oParagraph */
                $elementName = ($lvl == 0 ? "a:defRPr" : "a:lvl" . $lvl . "pPr");
                $objWriter->startElement($elementName);
                $objWriter->writeAttribute('algn', $oParagraph->getAlignment()->getHorizontal());
                $objWriter->writeAttribute('marL', CommonDrawing::pixelsToEmu($oParagraph->getAlignment()->getMarginLeft()));
                $objWriter->writeAttribute('marR', CommonDrawing::pixelsToEmu($oParagraph->getAlignment()->getMarginRight()));
                $objWriter->writeAttribute('indent', CommonDrawing::pixelsToEmu($oParagraph->getAlignment()->getIndent()));
                $objWriter->startElement('a:defRPr');
                $objWriter->writeAttribute('sz', $oParagraph->getFont()->getSize() * 100);
                $objWriter->writeAttribute('kern', '1200');
                $objWriter->startElement('a:solidFill');
                $objWriter->startElement('a:schemeClr');
                $objWriter->writeAttribute('val', $oParagraph->getFont()->getColor()->getValue());
                $objWriter->endElement();
                $objWriter->endElement();
                $objWriter->endElement();
                $objWriter->endElement();
            }
            $objWriter->endElement();
        }
        $objWriter->endElement();
        // > p:txStyles
        if (!is_null($pSlide->getTransition())) {
            $this->writeTransition($objWriter, $pSlide->getTransition());
        }
        $objWriter->endElement();
        // Return
        return $objWriter->getData();
    }
}