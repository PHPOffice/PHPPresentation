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
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\DocumentProperties;

class DocPropsCustom extends AbstractDecoratorWriter
{
    public function render(): ZipInterface
    {
        // Variables
        $pId = 0;

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Properties
        $objWriter->startElement('Properties');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties');
        $objWriter->writeAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

        if ($this->getPresentation()->getPresentationProperties()->isMarkedAsFinal()) {
            // property
            $objWriter->startElement('property');
            $objWriter->writeAttribute('fmtid', '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}');
            $objWriter->writeAttribute('pid', (++$pId) * 2);
            $objWriter->writeAttribute('name', '_MarkAsFinal');

            // property > vt:bool
            $objWriter->writeElement('vt:bool', 'true');

            // > property
            $objWriter->endElement();
        }

        $oDocumentProperties = $this->oPresentation->getDocumentProperties();
        foreach ($oDocumentProperties->getCustomProperties() as $customProperty) {
            $propertyValue = $oDocumentProperties->getCustomPropertyValue($customProperty);
            $propertyType = $oDocumentProperties->getCustomPropertyType($customProperty);

            $objWriter->startElement('property');
            $objWriter->writeAttribute('fmtid', '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}');
            $objWriter->writeAttribute('pid', (++$pId) * 2);
            $objWriter->writeAttribute('name', $customProperty);
            switch ($propertyType) {
                case DocumentProperties::PROPERTY_TYPE_INTEGER:
                    $objWriter->writeElement('vt:i4', (string) $propertyValue);

                    break;
                case DocumentProperties::PROPERTY_TYPE_FLOAT:
                    $objWriter->writeElement('vt:r8', (string) $propertyValue);

                    break;
                case DocumentProperties::PROPERTY_TYPE_BOOLEAN:
                    $objWriter->writeElement('vt:bool', $propertyValue ? 'true' : 'false');

                    break;
                case DocumentProperties::PROPERTY_TYPE_DATE:
                    $objWriter->startElement('vt:filetime');
                    $objWriter->writeRaw(date(DATE_W3C, (int) $propertyValue));
                    $objWriter->endElement();

                    break;
                default:
                    $objWriter->writeElement('vt:lpwstr', (string) $propertyValue);

                    break;
            }
            $objWriter->endElement();
        }

        // > Properties
        $objWriter->endElement();

        $this->oZip->addFromString('docProps/custom.xml', $objWriter->getData());

        // Return
        return $this->oZip;
    }
}
