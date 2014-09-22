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
use PhpOffice\PhpPowerpoint\Shape\Chart as ShapeChart;
use PhpOffice\PhpPowerpoint\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;
use PhpOffice\PhpPowerpoint\Shared\File;
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;

/**
 * \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\ContentTypes
 */
class ContentTypes extends AbstractPart
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
        $parentWriter = $this->getParentWriter();
        if (!$parentWriter instanceof PowerPoint2007) {
            throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007');
        }
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
        $masterSlides = $parentWriter->getLayoutPack()->getMasterSlides();
        foreach ($masterSlides as $masterSlide) {
            $this->writeOverrideContentType($objWriter, '/ppt/theme/theme' . $masterSlide['masterid'] . '.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
        }
            
        // Presentation
        $this->writeOverrideContentType($objWriter, '/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml');

        // PptProps
        $this->writeOverrideContentType($objWriter, '/ppt/presProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml');
        $this->writeOverrideContentType($objWriter, '/ppt/tableStyles.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml');
        $this->writeOverrideContentType($objWriter, '/ppt/viewProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml');
        
        // DocProps
        $this->writeOverrideContentType($objWriter, '/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
        $this->writeOverrideContentType($objWriter, '/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');
        
        // Slide masters
        $masterSlides = $parentWriter->getLayoutPack()->getMasterSlides();
        foreach ($masterSlides as $masterSlide) {
            $this->writeOverrideContentType($objWriter, '/ppt/slideMasters/slideMaster' . $masterSlide['masterid'] . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml');
        }

        // Slide layouts
        $slideLayouts = $parentWriter->getLayoutPack()->getLayouts();
        for ($i = 0; $i < count($slideLayouts); ++$i) {
            $this->writeOverrideContentType($objWriter, '/ppt/slideLayouts/slideLayout' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml');
        }

        // Slides
        $slideCount = $pPHPPowerPoint->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            $this->writeOverrideContentType($objWriter, '/ppt/slides/slide' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
        }

        // Add layoutpack content types
        $otherRelations = $parentWriter->getLayoutPack()->getMasterSlideRelations();
        foreach ($otherRelations as $otherRelation) {
            if (strpos($otherRelation['target'], 'http://') !== 0 && $otherRelation['contentType'] != '') {
                $this->writeOverrideContentType($objWriter, '/ppt/slideMasters/' . $otherRelation['target'], $otherRelation['contentType']);
            }
        }
        $otherRelations = $parentWriter->getLayoutPack()->getThemeRelations();
        foreach ($otherRelations as $otherRelation) {
            if (strpos($otherRelation['target'], 'http://') !== 0 && $otherRelation['contentType'] != '') {
                $this->writeOverrideContentType($objWriter, '/ppt/theme/' . $otherRelation['target'], $otherRelation['contentType']);
            }
        }
        $otherRelations = $parentWriter->getLayoutPack()->getLayoutRelations();
        foreach ($otherRelations as $otherRelation) {
            if (strpos($otherRelation['target'], 'http://') !== 0 && $otherRelation['contentType'] != '') {
                $this->writeOverrideContentType($objWriter, '/ppt/slideLayouts/' . $otherRelation['target'], $otherRelation['contentType']);
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
        $mediaCount = $parentWriter->getDrawingHashTable()->count();
        for ($i = 0; $i < $mediaCount; ++$i) {
            $extension = '';
            $mimeType  = '';

            $shapeIndex = $parentWriter->getDrawingHashTable()->getByIndex($i);
            if ($shapeIndex instanceof ShapeChart) {
                // Chart content type
                $this->writeOverrideContentType($objWriter, '/ppt/charts/chart' . $shapeIndex->getImageIndex() . '.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
            } else {
                if ($shapeIndex instanceof ShapeDrawing) {
                    $extension = strtolower($shapeIndex->getExtension());
                    $mimeType  = $this->getImageMimeType($shapeIndex->getPath());
                } elseif ($shapeIndex instanceof MemoryDrawing) {
                    $extension = strtolower($shapeIndex->getMimeType());
                    $extension = explode('/', $extension);
                    $extension = $extension[1];

                    $mimeType = $shapeIndex->getMimeType();
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
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter $objWriter    XML Writer
     * @param  string                         $pPartname    Part name
     * @param  string                         $pContentType Content type
     * @throws \Exception
     */
    private function writeDefaultContentType(XMLWriter $objWriter, $pPartname = '', $pContentType = '')
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
     * @param  \PhpOffice\PhpPowerpoint\Shared\XMLWriter $objWriter    XML Writer
     * @param  string                         $pPartname    Part name
     * @param  string                         $pContentType Content type
     * @throws \Exception
     */
    private function writeOverrideContentType(XMLWriter $objWriter, $pPartname = '', $pContentType = '')
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
