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
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Writer_PowerPoint2007_ContentTypes
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_ContentTypes extends PHPPowerPoint_Writer_PowerPoint2007_WriterPart
{
	/**
	 * Write content types to XML format
	 *
	 * @param 	PHPPowerPoint $pPHPPowerPoint
	 * @return 	string 						XML Output
	 * @throws 	Exception
	 */
	public function writeContentTypes(PHPPowerPoint $pPHPPowerPoint = null)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// Types
		$objWriter->startElement('Types');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

			// Rels
			$this->_writeDefaultContentType(
				$objWriter, 'rels', 'application/vnd.openxmlformats-package.relationships+xml'
			);

			// XML
			$this->_writeDefaultContentType(
				$objWriter, 'xml', 'application/xml'
			);

			// Themes
			$masterSlides = $this->getParentWriter()->getLayoutPack()->getMasterSlides();
			foreach ($masterSlides as $masterSlide) {
				$this->_writeOverrideContentType(
					$objWriter, '/ppt/theme/theme' . $masterSlide['masterid'] . '.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'
				);
			}
			
			// Presentation
			$this->_writeOverrideContentType(
				$objWriter, '/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
			);

			// DocProps
			$this->_writeOverrideContentType(
				$objWriter, '/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
			);

			$this->_writeOverrideContentType(
				$objWriter, '/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'
			);

			// Slide masters
			$masterSlides = $this->getParentWriter()->getLayoutPack()->getMasterSlides();
			foreach ($masterSlides as $masterSlide) {
				$this->_writeOverrideContentType(
					$objWriter, '/ppt/slideMasters/slideMaster' . $masterSlide['masterid'] . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
				);
			}

			// Slide layouts
			$slideLayouts = $this->getParentWriter()->getLayoutPack()->getLayouts();
			for ($i = 0; $i < count($slideLayouts); ++$i) {
				$this->_writeOverrideContentType(
					$objWriter, '/ppt/slideLayouts/slideLayout' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
				);
			}

			// Slides
			$slideCount = $pPHPPowerPoint->getSlideCount();
			for ($i = 0; $i < $slideCount; ++$i) {
				$this->_writeOverrideContentType(
					$objWriter, '/ppt/slides/slide' . ($i + 1) . '.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
				);
			}

			// Add layoutpack content types
			$otherRelations = null;
			$otherRelations = $this->getParentWriter()->getLayoutPack()->getMasterSlideRelations();
			foreach ($otherRelations as $otherRelations)
			{
				if (strpos($otherRelations['target'], 'http://') !== 0 && $otherRelations['contentType'] != '') {
    				$this->_writeOverrideContentType(
						$objWriter, '/ppt/slideMasters/' . $otherRelations['target'], $otherRelations['contentType']
					);
				}
			}
			$otherRelations = $this->getParentWriter()->getLayoutPack()->getThemeRelations();
			foreach ($otherRelations as $otherRelations)
			{
				if (strpos($otherRelations['target'], 'http://') !== 0 && $otherRelations['contentType'] != '') {
    				$this->_writeOverrideContentType(
						$objWriter, '/ppt/theme/' . $otherRelations['target'], $otherRelations['contentType']
					);
				}
			}
			$otherRelations = $this->getParentWriter()->getLayoutPack()->getLayoutRelations();
			foreach ($otherRelations as $otherRelations)
			{
				if (strpos($otherRelations['target'], 'http://') !== 0 && $otherRelations['contentType'] != '') {
    				$this->_writeOverrideContentType(
						$objWriter, '/ppt/slideLayouts/' . $otherRelations['target'], $otherRelations['contentType']
					);
				}
			}

			// Add media content-types
			$aMediaContentTypes = array();

			// GIF, JPEG, PNG
			$aMediaContentTypes['gif'] = 'image/gif';
			$aMediaContentTypes['jpg'] = 'image/jpeg';
			$aMediaContentTypes['jpeg'] = 'image/jpeg';
			$aMediaContentTypes['png'] = 'image/png';
			foreach ($aMediaContentTypes as $key => $value) {
				$this->_writeDefaultContentType($objWriter, $key, $value);
			}

			// Other media content types
			$mediaCount = $this->getParentWriter()->getDrawingHashTable()->count();
			for ($i = 0; $i < $mediaCount; ++$i) {
				$extension 	= '';
				$mimeType 	= '';

				if ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof PHPPowerPoint_Shape_Chart) {
    				$this->_writeOverrideContentType(
						$objWriter, '/ppt/charts/chart' . $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getImageIndex() . '.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
					);
				} else {
					if ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof PHPPowerPoint_Shape_Drawing) {
						$extension 	= strtolower($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getExtension());
						$mimeType 	= $this->_getImageMimeType( $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getPath() );
					} else if ($this->getParentWriter()->getDrawingHashTable()->getByIndex($i) instanceof PHPPowerPoint_Shape_MemoryDrawing) {
						$extension 	= strtolower($this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType());
						$extension 	= explode('/', $extension);
						$extension 	= $extension[1];
	
						$mimeType 	= $this->getParentWriter()->getDrawingHashTable()->getByIndex($i)->getMimeType();
					}
	
					if (!isset( $aMediaContentTypes[$extension]) ) {
							$aMediaContentTypes[$extension] = $mimeType;
	
							$this->_writeDefaultContentType(
								$objWriter, $extension, $mimeType
							);
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
	 * @param 	string	$pFile	Filename
	 * @return 	string	Mime Type
	 * @throws 	Exception
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

	/**
	 * Write Default content type
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	string 						$pPartname 		Part name
	 * @param 	string 						$pContentType 	Content type
	 * @throws 	Exception
	 */
	private function _writeDefaultContentType(PHPPowerPoint_Shared_XMLWriter $objWriter = null, $pPartname = '', $pContentType = '')
	{
		if ($pPartname != '' && $pContentType != '') {
			// Write content type
			$objWriter->startElement('Default');
			$objWriter->writeAttribute('Extension', 	$pPartname);
			$objWriter->writeAttribute('ContentType', 	$pContentType);
			$objWriter->endElement();
		} else {
			throw new Exception("Invalid parameters passed.");
		}
	}

	/**
	 * Write Override content type
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	string 						$pPartname 		Part name
	 * @param 	string 						$pContentType 	Content type
	 * @throws 	Exception
	 */
	private function _writeOverrideContentType(PHPPowerPoint_Shared_XMLWriter $objWriter = null, $pPartname = '', $pContentType = '')
	{
		if ($pPartname != '' && $pContentType != '') {
			// Write content type
			$objWriter->startElement('Override');
			$objWriter->writeAttribute('PartName', 		$pPartname);
			$objWriter->writeAttribute('ContentType', 	$pContentType);
			$objWriter->endElement();
		} else {
			throw new Exception("Invalid parameters passed.");
		}
	}
}
