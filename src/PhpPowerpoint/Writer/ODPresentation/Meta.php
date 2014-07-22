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

namespace PhpOffice\PhpPowerpoint\Writer\ODPresentation;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;

/**
 * \PhpOffice\PhpPowerpoint\Writer\ODPresentation\Meta
 */
class Meta extends AbstractPart
{
    /**
     * Write Meta file to XML format
     *
     * @param  PHPPowerPoint $pPHPPowerPoint
     * @return string        XML Output
     * @throws \Exception
     */
    public function writePart(PHPPowerPoint $pPHPPowerPoint)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8');

        // office:document-meta
        $objWriter->startElement('office:document-meta');
        $objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
        $objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
        $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $objWriter->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
        $objWriter->writeAttribute('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
        $objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
        $objWriter->writeAttribute('xmlns:smil', 'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0');
        $objWriter->writeAttribute('xmlns:anim', 'urn:oasis:names:tc:opendocument:xmlns:animation:1.0');
        $objWriter->writeAttribute('xmlns:grddl', 'http://www.w3.org/2003/g/data-view#');
        $objWriter->writeAttribute('xmlns:officeooo', 'http://openoffice.org/2009/office');
        $objWriter->writeAttribute('xmlns:drawooo', 'http://openoffice.org/2010/draw');
        $objWriter->writeAttribute('office:version', '1.2');

        // office:meta
        $objWriter->startElement('office:meta');

        // dc:creator
        $objWriter->writeElement('dc:creator', $pPHPPowerPoint->getProperties()->getLastModifiedBy());
        // dc:date
        $objWriter->writeElement('dc:date', gmdate('Y-m-d\TH:i:s.000', $pPHPPowerPoint->getProperties()->getModified()));
        // dc:description
        $objWriter->writeElement('dc:description', $pPHPPowerPoint->getProperties()->getDescription());
        // dc:subject
        $objWriter->writeElement('dc:subject', $pPHPPowerPoint->getProperties()->getSubject());
        // dc:title
        $objWriter->writeElement('dc:title', $pPHPPowerPoint->getProperties()->getTitle());
        // meta:creation-date
        $objWriter->writeElement('meta:creation-date', gmdate('Y-m-d\TH:i:s.000', $pPHPPowerPoint->getProperties()->getCreated()));
        // meta:initial-creator
        $objWriter->writeElement('meta:initial-creator', $pPHPPowerPoint->getProperties()->getCreator());
        // meta:keyword
        $objWriter->writeElement('meta:keyword', $pPHPPowerPoint->getProperties()->getKeywords());

        // @todo : Where these properties are written ?
        // $pPHPPowerPoint->getProperties()->getCategory()
        // $pPHPPowerPoint->getProperties()->getCompany()

        $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
