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
use PhpOffice\PhpPowerpoint\Shape\Chart;
use PhpOffice\PhpPowerpoint\Shape\Drawing;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;
use PhpOffice\PhpPowerpoint\Shared\File;
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;

/**
 * PHPPowerPoint_Writer_PowerPoint2007_ContentTypes
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class ContentTypes extends WriterPart
{
    /**
     * Write content types to XML format
     *
     * @param  PHPPowerPoint $pPHPPowerPoint
     * @return string        XML Output
     * @throws \Exception
     */
    public function writeContentTypes(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Types
        $objWriter->startElement('Types');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        // Rels
        $this->writeDefaultContentType($objWriter, 'rels', 'application/vnd.openxmlformats-package.relationships+xml');

        // XML
        $this->writeDefaultContentType($objWriter, 'xml', 'application/xml');

        // Themes
        $masterSlides = $this->getParentWriter()->getLayoutPack()->getMasterSlides();
        foreach ($masterSlides as $masterSlide) {
            $this->writeOverrideContentType($objWriter, '/ppt/theme/theme' . $masterSlide['masterid'] . '.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
        }

        // Presentation
        $this->writeOverrideContentType($objWriter, '/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml');

        // DocProps
        $this->writeOverrideContentType($objWriter, '/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');

        $this->writeOverrideContentType($objWriter, '/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');

        // Slide masters
        $masterSlides = $this->getParentWriter()->getLayoutPack()->getMasterSlides();
        foreach ($masterSlides as $masterSlide) {
            $this->writeOverrideContentType($objWriter, '/ppt/slideMasters/slideMaster' . $masterSlide['masterid'] . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml');
        }

        // Slide layouts
        $slideLayouts = $this->getParentWriter()->getLayoutPack()->getLayouts();
        for ($i = 0; $i < count($slideLayouts); ++$i) {
            $this->writeOverrideContentType($objWriter, '/ppt/slideLayouts/slideLayout' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml');
        }

        // Slides
        $slideCount = $pPHPPowerPoint->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            $this->writeOverrideContentType($objWriter, '/ppt/slides/slide' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
        }

        // Add layoutpack content types
        $otherRelations = $this->getParentWriter()->getLayoutPack()->getMasterSlideRelations();
        foreach ($otherRelations as $otherRelations) {
            if (strpos($otherRelations['target'], 'http://') !== 0 && $otherRelations['contentType'] != '') {
                $this->writeOverrideContentType($objWriter, '/ppt/slideMasters/' . $otherRelations['target'], $otherRelations['contentType']);
            }
        }
        $otherRelations = $this->getParentWriter()->getLayoutPack()->getThemeRelations();
        foreach ($otherRelations as $otherRelations) {
            if (strpos($otherRelations['target'], 'http://') !== 0 && $otherRelations['contentType'] != '') {
                $this->writeOverrideContentType($objWriter, '/ppt/theme/' . $otherRelations['target'], $otherRelations['contentType']);
            }
        }
        $otherRelations = $this->getParentWriter()->getLayoutPack()->getLayoutRelations();
        foreach ($otherRelations as $otherRelations) {
            if (strpos($otherRelations['target'], 'http://') !== 0 && $otherRelations['contentType'] != '') {
                $this->writeOverrideContentType($objWriter, '/ppt/slideLayouts/' . $otherRelations['target'], $otherRelations['contentType']);
            }
        }

        // Add media content-types
        $aMediaContentTypes = array();

        // GIF, JPEG, PNG
        $aMediaContentTypes['gif']  = 'image/gif';
        $aMediaContentTypes['jpg']  = 'image/jpeg';
        $aMediaContentTypes['jpeg'] = 'image/jpeg';
        $aMediaContentTypes['png']  = 'image/png';
        foreach ($aMediaContentTypes as $key => $value) {
            $this->writeDefaultContentType($objWriter, $key, $value);
        }

        // XLSX
        $this->writeDefaultContentType($objWriter, 'xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        // Other media content types
        $mediaCount = $this->getParentWriter()->getDrawingHashTable()->count();
        for ($i = 0; $i < $mediaCount; ++$i) {
            $extension = '';
            $mimeType  = '';

            if ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof Chart) {
                // Chart content type
                $this->writeOverrideContentType($objWriter, '/ppt/charts/chart' . $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getImageIndex() . '.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
            } else {
                if ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof Drawing) {
                    $extension = strtolower($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getExtension());
                    $mimeType  = $this->getImageMimeType($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath());
                } elseif ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof MemoryDrawing) {
                    $extension = strtolower($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType());
                    $extension = explode('/', $extension);
                    $extension = $extension[1];

                    $mimeType = $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType();
                }

                if (!isset($aMediaContentTypes[$extension])) {
                    $aMediaContentTypes[$extension] = $mimeType;

                    $this->writeDefaultContentType($objWriter, $extension, $mimeType);
                }
            }
        }

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Get image mime type
     *
     * @param  string    $pFile Filename
     * @return string    Mime Type
     * @throws \Exception
     */
    private function getImageMimeType($pFile = '')
    {
        if (File::fileExists($pFile)) {
            if (strpos($pFile, 'zip://') === 0) {
                $pZIPFile = str_replace('zip://', '', $pFile);
                $pZIPFile = substr($pZIPFile, 0, strpos($pZIPFile, '#'));
                $pImgFile = substr($pFile, strpos($pFile, '#') + 1);
                $oArchive = new \ZipArchive();
                $oArchive->open($pZIPFile);
                $image = getimagesizefromstring($oArchive->getFromName($pImgFile));
            } else {
                $image = getimagesize($pFile);
            }

            return image_type_to_mime_type($image[2]);
        } else {
            throw new \Exception("File $pFile does not exist");
        }
    }

    /**
     * Write Default content type
     *
     * @param  PHPPowerPoint_Shared_XMLWriter $objWriter    XML Writer
     * @param  string                         $pPartname    Part name
     * @param  string                         $pContentType Content type
     * @throws \Exception
     */
    private function writeDefaultContentType(XMLWriter $objWriter = null, $pPartname = '', $pContentType = '')
    {
        if ($pPartname != '' && $pContentType != '') {
            // Write content type
            $objWriter->startElement('Default');
            $objWriter->writeAttribute('Extension', $pPartname);
            $objWriter->writeAttribute('ContentType', $pContentType);
            $objWriter->endElement();
        } else {
            throw new \Exception("Invalid parameters passed.");
        }
    }

    /**
     * Write Override content type
     *
     * @param  PHPPowerPoint_Shared_XMLWriter $objWriter    XML Writer
     * @param  string                         $pPartname    Part name
     * @param  string                         $pContentType Content type
     * @throws \Exception
     */
    private function writeOverrideContentType(XMLWriter $objWriter = null, $pPartname = '', $pContentType = '')
    {
        if ($pPartname != '' && $pContentType != '') {
            // Write content type
            $objWriter->startElement('Override');
            $objWriter->writeAttribute('PartName', $pPartname);
            $objWriter->writeAttribute('ContentType', $pContentType);
            $objWriter->endElement();
        } else {
            throw new \Exception("Invalid parameters passed.");
        }
    }
}
