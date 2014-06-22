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

use PhpOffice\PhpPowerpoint\Writer\ODPresentation\AbstractPart;
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;
use PhpOffice\PhpPowerpoint\Shape\Drawing;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;
use PhpOffice\PhpPowerpoint\Shared\File;

/**
 * PHPPowerPoint_Writer_ODPresentation_Manifest
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Manifest extends AbstractPart
{
    /**
     * Write Manifest file to XML format
     *
     * @return string        XML Output
     * @throws \Exception
     */
    public function writeManifest()
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8');

        // manifest:manifest
        $objWriter->startElement('manifest:manifest');
        $objWriter->writeAttribute('xmlns:manifest', 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0');
        $objWriter->writeAttribute('manifest:version', '1.2');

        // manifest:file-entry
        $objWriter->startElement('manifest:file-entry');
        $objWriter->writeAttribute('manifest:media-type', 'application/vnd.oasis.opendocument.presentation');
        $objWriter->writeAttribute('manifest:version', '1.2');
        $objWriter->writeAttribute('manifest:full-path', '/');
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

        $arrMedia = array();
        for ($i = 0; $i < $this->getParentWriter()->getDrawingHashTable()->count(); ++$i) {
            if ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof Drawing) {
                if (!in_array(md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath()), $arrMedia)) {
                    $arrMedia[] = md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath());
                    $mimeType   = $this->getImageMimeType($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath());

                    $objWriter->startElement('manifest:file-entry');
                    $objWriter->writeAttribute('manifest:media-type', $mimeType);
                    $objWriter->writeAttribute('manifest:full-path', 'Pictures/' . md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath()) . '.' . $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getExtension());
                    $objWriter->endElement();
                }
            } elseif ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof MemoryDrawing) {
                if (!in_array(str_replace(' ', '_', $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getIndexedFilename()), $arrMedia)) {
                    $arrMedia[] = str_replace(' ', '_', $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getIndexedFilename());
                    $mimeType = $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType();

                    $objWriter->startElement('manifest:file-entry');
                    $objWriter->writeAttribute('manifest:media-type', $mimeType);
                    $objWriter->writeAttribute('manifest:full-path', 'Pictures/' . str_replace(' ', '_', $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getIndexedFilename()));
                    $objWriter->endElement();
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
}
