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

use PhpOffice\PhpPresentation\PhpPresentation;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007\DocProps
 */
class DocProps extends AbstractPart
{
    /**
     * Write docProps/app.xml to XML format
     *
     * @param  PhpPresentation $pPhpPresentation
     * @return string        XML Output
     * @throws \Exception
     */
    public function writeDocPropsApp(PhpPresentation $pPhpPresentation)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Properties
        $objWriter->startElement('Properties');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
        $objWriter->writeAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

        // Application
        $objWriter->writeElement('Application', 'Microsoft Office PowerPoint');

        // Slides
        $objWriter->writeElement('Slides', $pPhpPresentation->getSlideCount());

        // ScaleCrop
        $objWriter->writeElement('ScaleCrop', 'false');

        // HeadingPairs
        $objWriter->startElement('HeadingPairs');

        // Vector
        $objWriter->startElement('vt:vector');
        $objWriter->writeAttribute('size', '4');
        $objWriter->writeAttribute('baseType', 'variant');

        // Variant
        $objWriter->startElement('vt:variant');
        $objWriter->writeElement('vt:lpstr', 'Theme');
        $objWriter->endElement();

        // Variant
        $objWriter->startElement('vt:variant');
        $objWriter->writeElement('vt:i4', '1');
        $objWriter->endElement();

        // Variant
        $objWriter->startElement('vt:variant');
        $objWriter->writeElement('vt:lpstr', 'Slide Titles');
        $objWriter->endElement();

        // Variant
        $objWriter->startElement('vt:variant');
        $objWriter->writeElement('vt:i4', '1');
        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();

        // TitlesOfParts
        $objWriter->startElement('TitlesOfParts');

        // Vector
        $objWriter->startElement('vt:vector');
        $objWriter->writeAttribute('size', '1');
        $objWriter->writeAttribute('baseType', 'lpstr');

        $objWriter->writeElement('vt:lpstr', 'Office Theme');

        $objWriter->endElement();

        $objWriter->endElement();

        // Company
        $objWriter->writeElement('Company', $pPhpPresentation->getProperties()->getCompany());

        // LinksUpToDate
        $objWriter->writeElement('LinksUpToDate', 'false');

        // SharedDoc
        $objWriter->writeElement('SharedDoc', 'false');

        // HyperlinksChanged
        $objWriter->writeElement('HyperlinksChanged', 'false');

        // AppVersion
        $objWriter->writeElement('AppVersion', '12.0000');

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write docProps/core.xml to XML format
     *
     * @param  PhpPresentation $pPhpPresentation
     * @return string        XML Output
     * @throws \Exception
     */
    public function writeDocPropsCore(PhpPresentation $pPhpPresentation)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // cp:coreProperties
        $objWriter->startElement('cp:coreProperties');
        $objWriter->writeAttribute('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
        $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $objWriter->writeAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
        $objWriter->writeAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
        $objWriter->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

        // dc:creator
        $objWriter->writeElement('dc:creator', $pPhpPresentation->getProperties()->getCreator());

        // cp:lastModifiedBy
        $objWriter->writeElement('cp:lastModifiedBy', $pPhpPresentation->getProperties()->getLastModifiedBy());

        // dcterms:created
        $objWriter->startElement('dcterms:created');
        $objWriter->writeAttribute('xsi:type', 'dcterms:W3CDTF');
        $objWriter->writeRaw(gmdate("Y-m-d\TH:i:s\Z", $pPhpPresentation->getProperties()->getCreated()));
        $objWriter->endElement();

        // dcterms:modified
        $objWriter->startElement('dcterms:modified');
        $objWriter->writeAttribute('xsi:type', 'dcterms:W3CDTF');
        $objWriter->writeRaw(gmdate("Y-m-d\TH:i:s\Z", $pPhpPresentation->getProperties()->getModified()));
        $objWriter->endElement();

        // dc:title
        $objWriter->writeElement('dc:title', $pPhpPresentation->getProperties()->getTitle());

        // dc:description
        $objWriter->writeElement('dc:description', $pPhpPresentation->getProperties()->getDescription());

        // dc:subject
        $objWriter->writeElement('dc:subject', $pPhpPresentation->getProperties()->getSubject());

        // cp:keywords
        $objWriter->writeElement('cp:keywords', $pPhpPresentation->getProperties()->getKeywords());

        // cp:category
        $objWriter->writeElement('cp:category', $pPhpPresentation->getProperties()->getCategory());

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
