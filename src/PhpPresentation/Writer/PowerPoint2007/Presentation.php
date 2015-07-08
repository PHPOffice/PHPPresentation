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

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007\Workbook
 */
class Presentation extends AbstractPart
{
    /**
     * Write presentation to XML format
     *
     * @param  PhpPresentation $pPhpPresentation
     * @return string        XML Output
     * @throws \Exception
     */
    public function writePresentation(PhpPresentation $pPhpPresentation)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:presentation
        $objWriter->startElement('p:presentation');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        // p:sldMasterIdLst
        $objWriter->startElement('p:sldMasterIdLst');

        // Add slide masters
        $relationId    = 1;
        $slideMasterId = 2147483648;
        $parentWriter = $this->getParentWriter();
        if ($parentWriter instanceof PowerPoint2007) {
            $masterSlides  = $parentWriter->getLayoutPack()->getMasterSlides();
            $masterSlidesCount = count($masterSlides);
            // @todo foreach ($masterSlides as $masterSlide)
            for ($i = 0; $i < $masterSlidesCount; $i++) {
                // p:sldMasterId
                $objWriter->startElement('p:sldMasterId');
                $objWriter->writeAttribute('id', $slideMasterId);
                $objWriter->writeAttribute('r:id', 'rId' . $relationId++);
                $objWriter->endElement();
    
                // Increase identifier
                $slideMasterId += 12;
            }
        }
        $objWriter->endElement();

        // theme
        $relationId++;

        // p:sldIdLst
        $objWriter->startElement('p:sldIdLst');
        $this->writeSlides($objWriter, $pPhpPresentation, $relationId);
        $objWriter->endElement();

        // p:sldSz
        $objWriter->startElement('p:sldSz');
        $objWriter->writeAttribute('cx', $pPhpPresentation->getLayout()->getCX());
        $objWriter->writeAttribute('cy', $pPhpPresentation->getLayout()->getCY());
        if ($pPhpPresentation->getLayout()->getDocumentLayout() != DocumentLayout::LAYOUT_CUSTOM) {
            $objWriter->writeAttribute('type', $pPhpPresentation->getLayout()->getDocumentLayout());
        }
        $objWriter->endElement();

        // p:notesSz
        $objWriter->startElement('p:notesSz');
        $objWriter->writeAttribute('cx', '6858000');
        $objWriter->writeAttribute('cy', '9144000');
        $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write slides
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter       XML Writer
     * @param  PhpPresentation                  $pPhpPresentation
     * @param  int                            $startRelationId
     * @throws \Exception
     */
    private function writeSlides(XMLWriter $objWriter, PhpPresentation $pPhpPresentation, $startRelationId = 2)
    {
        // Write slides
        $slideCount = $pPhpPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            // p:sldId
            $this->writeSlide($objWriter, ($i + 256), ($i + $startRelationId));
        }
    }

    /**
     * Write slide
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  int                            $pSlideId  Slide id
     * @param  int                            $pRelId    Relationship ID
     * @throws \Exception
     */
    private function writeSlide(XMLWriter $objWriter, $pSlideId = 1, $pRelId = 1)
    {
        // p:sldId
        $objWriter->startElement('p:sldId');
        $objWriter->writeAttribute('id', $pSlideId);
        $objWriter->writeAttribute('r:id', 'rId' . $pRelId);
        $objWriter->endElement();
    }
}
