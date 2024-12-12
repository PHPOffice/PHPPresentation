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

class PptPresProps extends AbstractDecoratorWriter
{
    public function render(): ZipInterface
    {
        $presentationPpts = $this->oPresentation->getPresentationProperties();

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:presentationPr
        $objWriter->startElement('p:presentationPr');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // p:presentationPr > p:showPr
        $objWriter->startElement('p:showPr');
        if ($presentationPpts->isLoopContinuouslyUntilEsc()) {
            $objWriter->writeAttribute('loop', '1');
        }
        // Depends on the slideshow type
        // p:presentationPr > p:showPr > p:present
        // p:presentationPr > p:showPr > p:browse
        // p:presentationPr > p:showPr > p:kiosk
        $objWriter->writeElement('p:' . $presentationPpts->getSlideshowType());

        // > p:presentationPr > p:showPr
        $objWriter->endElement();

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

        $this->getZip()->addFromString('ppt/presProps.xml', $objWriter->getData());

        return $this->getZip();
    }
}
