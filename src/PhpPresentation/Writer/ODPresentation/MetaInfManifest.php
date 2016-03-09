<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\File;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
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
            if ($shape instanceof ShapeDrawing) {
                if (!in_array(md5($shape->getPath()), $arrMedia)) {
                    $arrMedia[] = md5($shape->getPath());
                    $mimeType   = $this->getImageMimeType($shape->getPath());

                    $objWriter->startElement('manifest:file-entry');
                    $objWriter->writeAttribute('manifest:media-type', $mimeType);
                    $objWriter->writeAttribute('manifest:full-path', 'Pictures/' . md5($shape->getPath()) . '.' . $shape->getExtension());
                    $objWriter->endElement();
                }
            } elseif ($shape instanceof MemoryDrawing) {
                if (!in_array(str_replace(' ', '_', $shape->getIndexedFilename()), $arrMedia)) {
                    $arrMedia[] = str_replace(' ', '_', $shape->getIndexedFilename());
                    $mimeType = $shape->getMimeType();

                    $objWriter->startElement('manifest:file-entry');
                    $objWriter->writeAttribute('manifest:media-type', $mimeType);
                    $objWriter->writeAttribute('manifest:full-path', 'Pictures/' . str_replace(' ', '_', $shape->getIndexedFilename()));
                    $objWriter->endElement();
                }
            }
        }

        foreach ($this->getPresentation()->getAllSlides() as $numSlide => $oSlide) {
            $oBkgImage = $oSlide->getBackground();
            if ($oBkgImage instanceof Image) {
                $mimeType   = $this->getImageMimeType($oBkgImage->getPath());

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

    /**
     * Get image mime type
     *
     * @param  string    $pFile Filename
     * @return string    Mime Type
     * @throws \Exception
     * @todo PowerPoint2007\ContentTypes duplicate Code : getImageMimeType
     */
    private function getImageMimeType($pFile = '')
    {
        if (strpos($pFile, 'zip://') === 0) {
            $pZIPFile = str_replace('zip://', '', $pFile);
            $pZIPFile = substr($pZIPFile, 0, strpos($pZIPFile, '#'));
            if (!File::fileExists($pZIPFile)) {
                throw new \Exception("File $pFile does not exist");
            }
            $pImgFile = substr($pFile, strpos($pFile, '#') + 1);
            $oArchive = new \ZipArchive();
            $oArchive->open($pZIPFile);
            if (!function_exists('getimagesizefromstring')) {
                $uri = 'data://application/octet-stream;base64,' . base64_encode($oArchive->getFromName($pImgFile));
                $image = getimagesize($uri);
            } else {
                $image = getimagesizefromstring($oArchive->getFromName($pImgFile));
            }
        } elseif (strpos($pFile, 'data:image/') === 0) {
            $sImage = $pFile;
            list(, $sImage) = explode(';', $sImage);
            list(, $sImage) = explode(',', $sImage);
            if (!function_exists('getimagesizefromstring')) {
                $uri = 'data://application/octet-stream;base64,' . base64_encode($sImage);
                $image = getimagesize($uri);
            } else {
                $image = getimagesizefromstring($sImage);
            }
        } else {
            if (!File::fileExists($pFile)) {
                throw new \Exception("File $pFile does not exist");
            }
            $image = getimagesize($pFile);
        }

        return image_type_to_mime_type($image[2]);
    }
}
