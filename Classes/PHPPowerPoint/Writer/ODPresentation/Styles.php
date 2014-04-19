<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */

/**
 * PHPPowerPoint_Writer_ODPresentation_Styles
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_ODPresentation_Styles extends PHPPowerPoint_Writer_ODPresentation_WriterPart
{
    /**
     * Write Meta file to XML format
     *
     * @param  PHPPowerPoint $pPHPPowerPoint
     * @return string        XML Output
     * @throws Exception
     */
    public function writeStyles(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
        }

        // XML header
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
        $stylePageLayout = $pPHPPowerPoint->getLayout()->getDocumentLayout();

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
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();

        // office:automatic-styles
        $objWriter->startElement('office:automatic-styles');
        // style:page-layout
        $objWriter->startElement('style:page-layout');
        if (empty($stylePageLayout)) {
            $objWriter->writeAttribute('style:name', 'sPL0');
        } else {
            $objWriter->writeAttribute('style:name', $stylePageLayout);
        }
        // style:page-layout-properties
        $objWriter->startElement('style:page-layout-properties');
        $objWriter->writeAttribute('fo:margin-top', '0cm');
        $objWriter->writeAttribute('fo:margin-bottom', '0cm');
        $objWriter->writeAttribute('fo:margin-left', '0cm');
        $objWriter->writeAttribute('fo:margin-right', '0cm');
        $objWriter->writeAttribute('fo:page-width', round(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters(PHPPowerPoint_Shared_Drawing::EMUToPixels($pPHPPowerPoint->getLayout()->getCX())), 1) . 'cm');
        $objWriter->writeAttribute('fo:page-height', round(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters(PHPPowerPoint_Shared_Drawing::EMUToPixels($pPHPPowerPoint->getLayout()->getCY())), 1) . 'cm');
        if ($pPHPPowerPoint->getLayout()->getCX() > $pPHPPowerPoint->getLayout()->getCY()) {
            $objWriter->writeAttribute('style:print-orientation', 'landscape');
        } else {
            $objWriter->writeAttribute('style:print-orientation', 'portrait');
        }
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();

        // office:master-styles
        $objWriter->startElement('office:master-styles');
        // style:master-page
        $objWriter->startElement('style:master-page');
        $objWriter->writeAttribute('style:name', 'Standard');
        $objWriter->writeAttribute('style:display-name', 'Standard');
        if (empty($stylePageLayout)) {
            $objWriter->writeAttribute('style:page-layout-name', 'sPL0');
        } else {
            $objWriter->writeAttribute('style:page-layout-name', $stylePageLayout);
        }
        $objWriter->writeAttribute('draw:style-name', 'sPres0');
        $objWriter->endElement();
        $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
