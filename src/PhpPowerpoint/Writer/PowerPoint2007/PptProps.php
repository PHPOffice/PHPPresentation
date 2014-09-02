<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;

class PptProps extends AbstractPart
{
    /**
     * Write ppt/presProps.xml to XML format
     *
     * @param     PhpPowerpoint    $pPHPPowerPoint
     * @return     string         XML Output
     * @throws     \Exception
     */
    public function writePresProps()
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:presentationPr
        $objWriter->startElement('p:presentationPr');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // p:extLst
        $objWriter->startElement('p:extLst');

        // p:ext
        $objWriter->startElement('p:ext');
        $objWriter->writeAttribute('uri', '{E76CE94A-603C-4142-B9EB-6D1370010A27}');

        // p14:discardImageEditData
        $objWriter->startElement('p14:discardImageEditData');
        $objWriter->writeAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
        $objWriter->writeAttribute('val', '0');
        $objWriter->endElement();

        // > p:ext
        $objWriter->endElement();

        // p:ext
        $objWriter->startElement('p:ext');
        $objWriter->writeAttribute('uri', '{D31A062A-798A-4329-ABDD-BBA856620510}');

        // p14:defaultImageDpi
        $objWriter->startElement('p14:defaultImageDpi');
        $objWriter->writeAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
        $objWriter->writeAttribute('val', '220');
        $objWriter->endElement();

        // > p:ext
        $objWriter->endElement();
        // > p:extLst
        $objWriter->endElement();
        // > p:presentationPr
        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * Write ppt/tableStyles.xml to XML format
     *
     * @param     PhpPowerpoint $pPHPPowerPoint
     * @return     string XML Output
     * @throws     \Exception
     */
    public function writeTableStyles()
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();
        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // a:tblStyleLst
        $objWriter->startElement('a:tblStyleLst');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('def', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
        $objWriter->endElement();

        return $objWriter->getData();
    }
    
    /**
     * Write ppt/viewProps.xml to XML format
     *
     * @param     PhpPowerpoint    $pPHPPowerPoint
     * @return     string         XML Output
     * @throws     \Exception
     */
    public function writeViewProps()
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:viewPr
        $objWriter->startElement('p:viewPr');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');
        $objWriter->writeAttribute('showComments', '0');

        // > p:viewPr
        $objWriter->endElement();

        return $objWriter->getData();
    }
}
