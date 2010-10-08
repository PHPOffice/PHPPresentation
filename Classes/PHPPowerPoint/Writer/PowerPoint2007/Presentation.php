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
 * PHPPowerPoint_Writer_PowerPoint2007_Workbook
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_Presentation extends PHPPowerPoint_Writer_PowerPoint2007_WriterPart
{
	/**
	 * Write presentation to XML format
	 *
	 * @param 	PHPPowerPoint	$pPHPPowerPoint
	 * @return 	string 		XML Output
	 * @throws 	Exception
	 */
	public function writePresentation(PHPPowerPoint $pPHPPowerPoint = null)
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

		// p:presentation
		$objWriter->startElement('p:presentation');
		$objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
		$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

			// p:sldMasterIdLst
			$objWriter->startElement('p:sldMasterIdLst');

				// Add slide masters
				$relationId = 1;
				$slideMasterId = 2147483648;
				$masterSlides = $this->getParentWriter()->getLayoutPack()->getMasterSlides();
				foreach ($masterSlides as $masterSlide) {
					// p:sldMasterId
					$objWriter->startElement('p:sldMasterId');
					$objWriter->writeAttribute('id',	$slideMasterId);
					$objWriter->writeAttribute('r:id',	'rId' . $relationId++);
					$objWriter->endElement();
					
					// Increase identifier
					$slideMasterId += 12;
				}

			$objWriter->endElement();

			// theme
			$relationId++;
			
			// p:sldIdLst
			$objWriter->startElement('p:sldIdLst');
			$this->_writeSlides($objWriter, $pPHPPowerPoint, $relationId);
			$objWriter->endElement();

			// p:sldSz
			$objWriter->startElement('p:sldSz');
			$objWriter->writeAttribute('cx', '9144000');
			$objWriter->writeAttribute('cy', '6858000');
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
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPPowerPoint					$pPHPPowerPoint
	 * @param	int								$startRelationId
	 * @throws 	Exception
	 */
	private function _writeSlides(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint $pPHPPowerPoint = null, $startRelationId = 2)
	{
		// Write slides
		$slideCount = $pPHPPowerPoint->getSlideCount();
		for ($i = 0; $i < $slideCount; ++$i) {
			// p:sldId
			$this->_writeSlide(
				$objWriter,
				($i + 256),
				($i + $startRelationId)
			);
		}
	}

	/**
	 * Write slide
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	int							$pSlideId	 		Slide id
	 * @param 	int							$pRelId				Relationship ID
	 * @throws 	Exception
	 */
	private function _writeSlide(PHPPowerPoint_Shared_XMLWriter $objWriter = null, $pSlideId = 1, $pRelId = 1)
	{
		// p:sldId
		$objWriter->startElement('p:sldId');
		$objWriter->writeAttribute('id', 	$pSlideId);
		$objWriter->writeAttribute('r:id', 	'rId' . $pRelId);
		$objWriter->endElement();
	}
}
