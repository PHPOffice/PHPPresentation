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

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\DocumentProperties;

class Meta extends AbstractDecoratorWriter
{
    public function render(): ZipInterface
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
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
        $objWriter->writeElement('dc:creator', $this->getPresentation()->getDocumentProperties()->getLastModifiedBy());
        // dc:date
        $objWriter->writeElement('dc:date', gmdate('Y-m-d\TH:i:s.000', $this->getPresentation()->getDocumentProperties()->getModified()));
        // dc:description
        $objWriter->writeElement('dc:description', $this->getPresentation()->getDocumentProperties()->getDescription());
        // dc:subject
        $objWriter->writeElement('dc:subject', $this->getPresentation()->getDocumentProperties()->getSubject());
        // dc:title
        $objWriter->writeElement('dc:title', $this->getPresentation()->getDocumentProperties()->getTitle());
        // meta:creation-date
        $objWriter->writeElement('meta:creation-date', gmdate('Y-m-d\TH:i:s.000', $this->getPresentation()->getDocumentProperties()->getCreated()));
        // meta:initial-creator
        $objWriter->writeElement('meta:initial-creator', $this->getPresentation()->getDocumentProperties()->getCreator());
        // meta:keyword
        $objWriter->writeElement('meta:keyword', $this->getPresentation()->getDocumentProperties()->getKeywords());
        // meta:generator
        $objWriter->writeElement('meta:generator', $this->getPresentation()->getDocumentProperties()->getGenerator());

        // meta:user-defined
        $oDocumentProperties = $this->oPresentation->getDocumentProperties();
        foreach ($oDocumentProperties->getCustomProperties() as $customProperty) {
            $propertyValue = $oDocumentProperties->getCustomPropertyValue($customProperty);
            $propertyType = $oDocumentProperties->getCustomPropertyType($customProperty);

            $objWriter->startElement('meta:user-defined');
            $objWriter->writeAttribute('meta:name', $customProperty);
            switch ($propertyType) {
                case DocumentProperties::PROPERTY_TYPE_INTEGER:
                case DocumentProperties::PROPERTY_TYPE_FLOAT:
                    $objWriter->writeAttribute('meta:value-type', 'float');
                    $objWriter->writeRaw((string) $propertyValue);

                    break;
                case DocumentProperties::PROPERTY_TYPE_BOOLEAN:
                    $objWriter->writeAttribute('meta:value-type', 'boolean');
                    $objWriter->writeRaw($propertyValue ? 'true' : 'false');

                    break;
                case DocumentProperties::PROPERTY_TYPE_DATE:
                    $objWriter->writeAttribute('meta:value-type', 'date');
                    $objWriter->writeRaw(date(DATE_W3C, (int) $propertyValue));

                    break;
                case DocumentProperties::PROPERTY_TYPE_STRING:
                case DocumentProperties::PROPERTY_TYPE_UNKNOWN:
                default:
                    $objWriter->writeAttribute('meta:value-type', 'string');
                    $objWriter->writeRaw((string) $propertyValue);

                    break;
            }
            $objWriter->endElement();
        }

        // @todo : Where these properties are written ?
        // $this->getPresentation()->getDocumentProperties()->getCategory()
        // $this->getPresentation()->getDocumentProperties()->getCompany()

        $objWriter->endElement();

        $objWriter->endElement();

        $this->getZip()->addFromString('meta.xml', $objWriter->getData());

        return $this->getZip();
    }
}
