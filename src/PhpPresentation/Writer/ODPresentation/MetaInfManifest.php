<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Writer\ODPresentation;

class MetaInfManifest extends AbstractDecoratorWriter
{
    /**
     * @return ZipInterface
     */
    public function render()
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8');

        // manifest:manifest
        $objWriter->startElement('manifest:manifest');
        $objWriter->writeAttribute('xmlns:manifest', 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0');
        $objWriter->writeAttribute('manifest:version', '1.2');

        // manifest:file-entry
        $objWriter->startElement('manifest:file-entry');
        $objWriter->writeAttribute('manifest:media-type', 'application/vnd.oasis.opendocument.presentation');
        $objWriter->writeAttribute('manifest:full-path', '/');
        $objWriter->writeAttribute('manifest:version', '1.2');
        $objWriter->endElement();
        // manifest:file-entry
        $objWriter->startElement('manifest:file-entry');
        $objWriter->writeAttribute('manifest:media-type', 'text/xml');
        $objWriter->writeAttribute('manifest:full-path', 'content.xml');
        $objWriter->endElement();
        // manifest:file-entry
        $objWriter->startElement('manifest:file-entry');
        $objWriter->writeAttribute('manifest:media-type', 'text/xml');
        $objWriter->writeAttribute('manifest:full-path', 'meta.xml');
        $objWriter->endElement();
        // manifest:file-entry
        $objWriter->startElement('manifest:file-entry');
        $objWriter->writeAttribute('manifest:media-type', 'text/xml');
        $objWriter->writeAttribute('manifest:full-path', 'styles.xml');
        $objWriter->endElement();

        // Charts
        foreach ($this->getArrayChart() as $key => $shape) {
            $objWriter->startElement('manifest:file-entry');
            $objWriter->writeAttribute('manifest:media-type', 'application/vnd.oasis.opendocument.chart');
            $objWriter->writeAttribute('manifest:full-path', 'Object '.$key.'/');
            $objWriter->endElement();
            $objWriter->startElement('manifest:file-entry');
            $objWriter->writeAttribute('manifest:media-type', 'text/xml');
            $objWriter->writeAttribute('manifest:full-path', 'Object '.$key.'/content.xml');
            $objWriter->endElement();
        }

        $arrMedia = array();
        for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
            $shape = $this->getDrawingHashTable()->getByIndex($i);
            if (! ($shape instanceof ShapeDrawing\AbstractDrawingAdapter)) {
                continue;
            }
            $arrMedia[] = $shape->getIndexedFilename();
            $objWriter->startElement('manifest:file-entry');
            $objWriter->writeAttribute('manifest:media-type', $shape->getMimeType());
            $objWriter->writeAttribute('manifest:full-path', 'Pictures/' . $shape->getIndexedFilename());
            $objWriter->endElement();
        }

        foreach ($this->getPresentation()->getAllSlides() as $numSlide => $oSlide) {
            $oBkgImage = $oSlide->getBackground();
            if ($oBkgImage instanceof Image) {
                $arrayImage = getimagesize($oBkgImage->getPath());
                $mimeType  = image_type_to_mime_type($arrayImage[2]);

                $objWriter->startElement('manifest:file-entry');
                $objWriter->writeAttribute('manifest:media-type', $mimeType);
                $objWriter->writeAttribute('manifest:full-path', 'Pictures/' . str_replace(' ', '_', $oBkgImage->getIndexedFilename($numSlide)));
                $objWriter->endElement();
            }
        }

        if ($this->getPresentation()->getPresentationProperties()->getThumbnailPath()) {
            $pathThumbnail = $this->getPresentation()->getPresentationProperties()->getThumbnailPath();
            // Size : 128x128 pixel
            // PNG : 8bit, non-interlaced with full alpha transparency
            $gdImage = imagecreatefromstring(file_get_contents($pathThumbnail));
            if ($gdImage) {
                imagedestroy($gdImage);
                $objWriter->startElement('manifest:file-entry');
                $objWriter->writeAttribute('manifest:media-type', 'image/png');
                $objWriter->writeAttribute('manifest:full-path', 'Thumbnails/thumbnail.png');
                $objWriter->endElement();
            }
        }

        $objWriter->endElement();

        $this->getZip()->addFromString('META-INF/manifest.xml', $objWriter->getData());
        return $this->getZip();
    }
}
