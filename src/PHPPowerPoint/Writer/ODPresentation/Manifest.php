<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */

/**
 * PHPPowerPoint_Writer_ODPresentation_Manifest
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_ODPresentation_Manifest extends PHPPowerPoint_Writer_ODPresentation_WriterPart
{
    /**
     * Write Manifest file to XML format
     *
     * @return string        XML Output
     * @throws Exception
     */
    public function writeManifest()
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
        }

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
            if ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof PHPPowerPoint_Shape_Drawing) {
                if (!in_array(md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath()), $arrMedia)) {
                    $arrMedia[] = md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath());
                    $extension  = strtolower($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getExtension());
                    $mimeType   = $this->_getImageMimeType($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath());

                    $objWriter->startElement('manifest:file-entry');
                    $objWriter->writeAttribute('manifest:media-type', $mimeType);
                    $objWriter->writeAttribute('manifest:full-path', 'Pictures/' . md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath()) . '.' . $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getExtension());
                    $objWriter->endElement();
                }
            } elseif ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof PHPPowerPoint_Shape_MemoryDrawing) {
                if (!in_array(md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath()), $arrMedia)) {
                    $arrMedia[] = md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath());

                    $extension = strtolower($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType());
                    $extension = explode('/', $extension);
                    $extension = $extension[1];

                    $mimeType = $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType();

                    $objWriter->startElement('manifest:file-entry');
                    $objWriter->writeAttribute('manifest:media-type', $mimeType);
                    $objWriter->writeAttribute('manifest:full-path', 'Pictures/' . md5($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath()) . '.' . $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getExtension());
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
     * @throws Exception
     */
    private function _getImageMimeType($pFile = '')
    {
        if (PHPPowerPoint_Shared_File::file_exists($pFile)) {
            $image = getimagesize($pFile);

            return image_type_to_mime_type($image[2]);
        } else {
            throw new Exception("File $pFile does not exist");
        }
    }
}
