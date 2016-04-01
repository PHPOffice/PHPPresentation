<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Border;

class Styles extends AbstractDecoratorWriter
{
    /**
     * Stores font styles draw:gradient nodes
     *
     * @var array
     */
    protected $arrayGradient = array();
    /**
     * Stores font styles draw:stroke-dash nodes
     *
     * @var array
     */
    protected $arrayStrokeDash = array();

    /**
     * @return ZipInterface
     */
    public function render()
    {
        $this->getZip()->addFromString('styles.xml', $this->writePart());
        return $this->getZip();
    }

    /**
     * Write Meta file to XML format
     *
     * @return string        XML Output
     * @throws \Exception
     */
    protected function writePart()
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        $objWriter->startDocument('1.0', 'UTF-8');

        // office:document-meta
        $objWriter->startElement('office:document-styles');
        $objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
        $objWriter->writeAttribute('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0');
        $objWriter->writeAttribute('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
        $objWriter->writeAttribute('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
        $objWriter->writeAttribute('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
        $objWriter->writeAttribute('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0');
        $objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
        $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $objWriter->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
        $objWriter->writeAttribute('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0');
        $objWriter->writeAttribute('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
        $objWriter->writeAttribute('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0');
        $objWriter->writeAttribute('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0');
        $objWriter->writeAttribute('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0');
        $objWriter->writeAttribute('xmlns:math', 'http://www.w3.org/1998/Math/MathML');
        $objWriter->writeAttribute('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0');
        $objWriter->writeAttribute('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0');
        $objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
        $objWriter->writeAttribute('xmlns:ooow', 'http://openoffice.org/2004/writer');
        $objWriter->writeAttribute('xmlns:oooc', 'http://openoffice.org/2004/calc');
        $objWriter->writeAttribute('xmlns:dom', 'http://www.w3.org/2001/xml-events');
        $objWriter->writeAttribute('xmlns:smil', 'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0');
        $objWriter->writeAttribute('xmlns:anim', 'urn:oasis:names:tc:opendocument:xmlns:animation:1.0');
        $objWriter->writeAttribute('xmlns:rpt', 'http://openoffice.org/2005/report');
        $objWriter->writeAttribute('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2');
        $objWriter->writeAttribute('xmlns:xhtml', 'http://www.w3.org/1999/xhtml');
        $objWriter->writeAttribute('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#');
        $objWriter->writeAttribute('xmlns:officeooo', 'http://openoffice.org/2009/office');
        $objWriter->writeAttribute('xmlns:tableooo', 'http://openoffice.org/2009/table');
        $objWriter->writeAttribute('xmlns:drawooo', 'http://openoffice.org/2010/draw');
        $objWriter->writeAttribute('xmlns:css3t', 'http://www.w3.org/TR/css3-text/');
        $objWriter->writeAttribute('office:version', '1.2');

        // Variables
        $stylePageLayout = $this->getPresentation()->getLayout()->getDocumentLayout();
        if (empty($stylePageLayout)) {
            $stylePageLayout = 'sPL0';
        }

        // office:styles
        $objWriter->startElement('office:styles');
        // style:style
        $objWriter->startElement('style:style');
        $objWriter->writeAttribute('style:name', 'sPres0');
        $objWriter->writeAttribute('style:display-name', 'sPres0');
        $objWriter->writeAttribute('style:family', 'presentation');
        // style:graphic-properties
        $objWriter->startElement('style:graphic-properties');
        $objWriter->writeAttribute('draw:fill-color', '#ffffff');
        // > style:graphic-properties
        $objWriter->endElement();
        // > style:style
        $objWriter->endElement();

        foreach ($this->getPresentation()->getAllSlides() as $keySlide => $oSlide) {
            foreach ($oSlide->getShapeCollection() as $shape) {
                if ($shape instanceof Table) {
                    $this->writeTableStyle($objWriter, $shape);
                } elseif ($shape instanceof Group) {
                    $this->writeGroupStyle($objWriter, $shape);
                } elseif ($shape instanceof RichText) {
                    $this->writeRichTextStyle($objWriter, $shape);
                }
            }
            $oBkgImage = $oSlide->getBackground();
            if ($oBkgImage instanceof Image) {
                $this->writeBackgroundStyle($objWriter, $oBkgImage, $keySlide);
            }
        }
        // > office:styles
        $objWriter->endElement();

        // office:automatic-styles
        $objWriter->startElement('office:automatic-styles');
        // style:page-layout
        $objWriter->startElement('style:page-layout');
        $objWriter->writeAttribute('style:name', $stylePageLayout);
        // style:page-layout-properties
        $objWriter->startElement('style:page-layout-properties');
        $objWriter->writeAttribute('fo:margin-top', '0cm');
        $objWriter->writeAttribute('fo:margin-bottom', '0cm');
        $objWriter->writeAttribute('fo:margin-left', '0cm');
        $objWriter->writeAttribute('fo:margin-right', '0cm');
        $objWriter->writeAttribute('fo:page-width', Text::numberFormat(CommonDrawing::pixelsToCentimeters(CommonDrawing::emuToPixels($this->getPresentation()->getLayout()->getCX())), 1) . 'cm');
        $objWriter->writeAttribute('fo:page-height', Text::numberFormat(CommonDrawing::pixelsToCentimeters(CommonDrawing::emuToPixels($this->getPresentation()->getLayout()->getCY())), 1) . 'cm');
        $printOrientation = 'portrait';
        if ($this->getPresentation()->getLayout()->getCX() > $this->getPresentation()->getLayout()->getCY()) {
            $printOrientation = 'landscape';
        }
        $objWriter->writeAttribute('style:print-orientation', $printOrientation);
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();

        // office:master-styles
        $objWriter->startElement('office:master-styles');
        // style:master-page
        $objWriter->startElement('style:master-page');
        $objWriter->writeAttribute('style:name', 'Standard');
        $objWriter->writeAttribute('style:display-name', 'Standard');
        $objWriter->writeAttribute('style:page-layout-name', $stylePageLayout);
        $objWriter->writeAttribute('draw:style-name', 'sPres0');
        $objWriter->endElement();
        $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write the default style information for a RichText shape
     *
     * @param XMLWriter $objWriter
     * @param RichText $shape
     */
    protected function writeRichTextStyle(XMLWriter $objWriter, RichText $shape)
    {
        $oFill = $shape->getFill();
        if ($oFill->getFillType() == Fill::FILL_GRADIENT_LINEAR || $oFill->getFillType() == Fill::FILL_GRADIENT_PATH) {
            if (!in_array($oFill->getHashCode(), $this->arrayGradient)) {
                $this->writeGradientFill($objWriter, $oFill);
            }
        }
        $oBorder = $shape->getBorder();
        if ($oBorder->getDashStyle() != Border::DASH_SOLID) {
            if (!in_array($oBorder->getDashStyle(), $this->arrayStrokeDash)) {
                $objWriter->startElement('draw:stroke-dash');
                $objWriter->writeAttribute('draw:name', 'strokeDash_'.$oBorder->getDashStyle());
                $objWriter->writeAttribute('draw:style', 'rect');
                switch ($oBorder->getDashStyle()) {
                    case Border::DASH_DASH:
                        $objWriter->writeAttribute('draw:distance', '0.105cm');
                        $objWriter->writeAttribute('draw:dots2', '1');
                        $objWriter->writeAttribute('draw:dots2-length', '0.14cm');
                        break;
                    case Border::DASH_DASHDOT:
                        $objWriter->writeAttribute('draw:distance', '0.105cm');
                        $objWriter->writeAttribute('draw:dots1', '1');
                        $objWriter->writeAttribute('draw:dots1-length', '0.035cm');
                        $objWriter->writeAttribute('draw:dots2', '1');
                        $objWriter->writeAttribute('draw:dots2-length', '0.14cm');
                        break;
                    case Border::DASH_DOT:
                        $objWriter->writeAttribute('draw:distance', '0.105cm');
                        $objWriter->writeAttribute('draw:dots1', '1');
                        $objWriter->writeAttribute('draw:dots1-length', '0.035cm');
                        break;
                    case Border::DASH_LARGEDASH:
                        $objWriter->writeAttribute('draw:distance', '0.105cm');
                        $objWriter->writeAttribute('draw:dots2', '1');
                        $objWriter->writeAttribute('draw:dots2-length', '0.28cm');
                        break;
                    case Border::DASH_LARGEDASHDOT:
                        $objWriter->writeAttribute('draw:distance', '0.105cm');
                        $objWriter->writeAttribute('draw:dots1', '1');
                        $objWriter->writeAttribute('draw:dots1-length', '0.035cm');
                        $objWriter->writeAttribute('draw:dots2', '1');
                        $objWriter->writeAttribute('draw:dots2-length', '0.28cm');
                        break;
                    case Border::DASH_LARGEDASHDOTDOT:
                        $objWriter->writeAttribute('draw:distance', '0.105cm');
                        $objWriter->writeAttribute('draw:dots1', '2');
                        $objWriter->writeAttribute('draw:dots1-length', '0.035cm');
                        $objWriter->writeAttribute('draw:dots2', '1');
                        $objWriter->writeAttribute('draw:dots2-length', '0.28cm');
                        break;
                    case Border::DASH_SYSDASH:
                        $objWriter->writeAttribute('draw:distance', '0.035cm');
                        $objWriter->writeAttribute('draw:dots2', '1');
                        $objWriter->writeAttribute('draw:dots2-length', '0.105cm');
                        break;
                    case Border::DASH_SYSDASHDOT:
                        $objWriter->writeAttribute('draw:distance', '0.035cm');
                        $objWriter->writeAttribute('draw:dots1', '1');
                        $objWriter->writeAttribute('draw:dots1-length', '0.035cm');
                        $objWriter->writeAttribute('draw:dots2', '1');
                        $objWriter->writeAttribute('draw:dots2-length', '0.105cm');
                        break;
                    case Border::DASH_SYSDASHDOTDOT:
                        $objWriter->writeAttribute('draw:distance', '0.035cm');
                        $objWriter->writeAttribute('draw:dots1', '2');
                        $objWriter->writeAttribute('draw:dots1-length', '0.035cm');
                        $objWriter->writeAttribute('draw:dots2', '1');
                        $objWriter->writeAttribute('draw:dots2-length', '0.105cm');
                        break;
                    case Border::DASH_SYSDOT:
                        $objWriter->writeAttribute('draw:distance', '0.035cm');
                        $objWriter->writeAttribute('draw:dots1', '1');
                        $objWriter->writeAttribute('draw:dots1-length', '0.035cm');
                        break;
                }
                $objWriter->endElement();
                $this->arrayStrokeDash[] = $oBorder->getDashStyle();
            }
        }
    }

    /**
     * Write the default style information for a Table shape
     *
     * @param XMLWriter $objWriter
     * @param Table $shape
     */
    protected function writeTableStyle(XMLWriter $objWriter, Table $shape)
    {
        foreach ($shape->getRows() as $row) {
            foreach ($row->getCells() as $cell) {
                if ($cell->getFill()->getFillType() == Fill::FILL_GRADIENT_LINEAR) {
                    if (!in_array($cell->getFill()->getHashCode(), $this->arrayGradient)) {
                        $this->writeGradientFill($objWriter, $cell->getFill());
                    }
                }
            }
        }
    }

    /**
     * Writes the style information for a group of shapes
     *
     * @param XMLWriter $objWriter
     * @param Group $group
     */
    protected function writeGroupStyle(XMLWriter $objWriter, Group $group)
    {
        $shapes = $group->getShapeCollection();
        foreach ($shapes as $shape) {
            if ($shape instanceof Table) {
                $this->writeTableStyle($objWriter, $shape);
            } elseif ($shape instanceof Group) {
                $this->writeGroupStyle($objWriter, $shape);
            }
        }
    }

    /**
     * Write the gradient style
     * @param XMLWriter $objWriter
     * @param Fill $oFill
     */
    protected function writeGradientFill(XMLWriter $objWriter, Fill $oFill)
    {
        $objWriter->startElement('draw:gradient');
        $objWriter->writeAttribute('draw:name', 'gradient_'.$oFill->getHashCode());
        $objWriter->writeAttribute('draw:display-name', 'gradient_'.$oFill->getHashCode());
        $objWriter->writeAttribute('draw:style', 'linear');
        $objWriter->writeAttribute('draw:start-intensity', '100%');
        $objWriter->writeAttribute('draw:end-intensity', '100%');
        $objWriter->writeAttribute('draw:start-color', '#'.$oFill->getStartColor()->getRGB());
        $objWriter->writeAttribute('draw:end-color', '#'.$oFill->getEndColor()->getRGB());
        $objWriter->writeAttribute('draw:border', '0%');
        $objWriter->writeAttribute('draw:angle', $oFill->getRotation() - 90);
        $objWriter->endElement();
        $this->arrayGradient[] = $oFill->getHashCode();
    }

    /**
     * Write the background image style
     * @param XMLWriter $objWriter
     * @param Image $oBkgImage
     */
    protected function writeBackgroundStyle(XMLWriter $objWriter, Image $oBkgImage, $numSlide)
    {
        $objWriter->startElement('draw:fill-image');
        $objWriter->writeAttribute('draw:name', 'background_'.$numSlide);
        $objWriter->writeAttribute('xlink:href', 'Pictures/'.str_replace(' ', '_', $oBkgImage->getIndexedFilename($numSlide)));
        $objWriter->writeAttribute('xlink:type', 'simple');
        $objWriter->writeAttribute('xlink:show', 'embed');
        $objWriter->writeAttribute('xlink:actuate', 'onLoad');
        $objWriter->endElement();
    }
}
