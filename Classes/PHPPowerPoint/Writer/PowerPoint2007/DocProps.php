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
 * PHPPowerPoint_Writer_PowerPoint2007_DocProps
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_DocProps extends PHPPowerPoint_Writer_PowerPoint2007_WriterPart
{
/**
	 * Write docProps/app.xml to XML format
	 *
	 * @param 	PHPPowerPoint	$pPHPPowerPoint
	 * @return 	string 		XML Output
	 * @throws 	Exception
	 */
	public function writeDocPropsApp(PHPPowerPoint $pPHPPowerPoint = null)
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

		// Properties
		$objWriter->startElement('Properties');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
		$objWriter->writeAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

			// Application
			$objWriter->writeElement('Application', 	'Microsoft Office PowerPoint');

			// Slides
			$objWriter->writeElement('Slides', 	$pPHPPowerPoint->getSlideCount());

			// ScaleCrop
			$objWriter->writeElement('ScaleCrop', 		'false');

			// HeadingPairs
			$objWriter->startElement('HeadingPairs');

				// Vector
				$objWriter->startElement('vt:vector');
				$objWriter->writeAttribute('size', 		'4');
				$objWriter->writeAttribute('baseType', 	'variant');

					// Variant
					$objWriter->startElement('vt:variant');
						$objWriter->writeElement('vt:lpstr', 	'Theme');
					$objWriter->endElement();

					// Variant
					$objWriter->startElement('vt:variant');
						$objWriter->writeElement('vt:i4', 		'1');
					$objWriter->endElement();

					// Variant
					$objWriter->startElement('vt:variant');
						$objWriter->writeElement('vt:lpstr', 	'Slide Titles');
					$objWriter->endElement();

					// Variant
					$objWriter->startElement('vt:variant');
						$objWriter->writeElement('vt:i4', 		'1');
					$objWriter->endElement();

				$objWriter->endElement();

			$objWriter->endElement();

			// TitlesOfParts
			$objWriter->startElement('TitlesOfParts');

				// Vector
				$objWriter->startElement('vt:vector');
				$objWriter->writeAttribute('size', 		'1');
				$objWriter->writeAttribute('baseType',	'lpstr');

					$objWriter->writeElement('vt:lpstr', 	'Office Theme');

				$objWriter->endElement();

			$objWriter->endElement();

			// Company
			$objWriter->writeElement('Company', 			$pPHPPowerPoint->getProperties()->getCompany());

			// LinksUpToDate
			$objWriter->writeElement('LinksUpToDate', 		'false');

			// SharedDoc
			$objWriter->writeElement('SharedDoc', 			'false');

			// HyperlinksChanged
			$objWriter->writeElement('HyperlinksChanged', 	'false');

			// AppVersion
			$objWriter->writeElement('AppVersion', 			'12.0000');

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write docProps/core.xml to XML format
	 *
	 * @param 	PHPPowerPoint	$pPHPPowerPoint
	 * @return 	string 		XML Output
	 * @throws 	Exception
	 */
	public function writeDocPropsCore(PHPPowerPoint $pPHPPowerPoint = null)
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

		// cp:coreProperties
		$objWriter->startElement('cp:coreProperties');
		$objWriter->writeAttribute('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
		$objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
		$objWriter->writeAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
		$objWriter->writeAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
		$objWriter->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

			// dc:creator
			$objWriter->writeElement('dc:creator',			$pPHPPowerPoint->getProperties()->getCreator());

			// cp:lastModifiedBy
			$objWriter->writeElement('cp:lastModifiedBy', 	$pPHPPowerPoint->getProperties()->getLastModifiedBy());

			// dcterms:created
			$objWriter->startElement('dcterms:created');
			$objWriter->writeAttribute('xsi:type', 'dcterms:W3CDTF');
			$objWriter->writeRaw(gmdate("Y-m-d\TH:i:s\Z", 	$pPHPPowerPoint->getProperties()->getCreated()));
			$objWriter->endElement();

			// dcterms:modified
			$objWriter->startElement('dcterms:modified');
			$objWriter->writeAttribute('xsi:type', 'dcterms:W3CDTF');
			$objWriter->writeRaw(gmdate("Y-m-d\TH:i:s\Z", 	$pPHPPowerPoint->getProperties()->getModified()));
			$objWriter->endElement();

			// dc:title
			$objWriter->writeElement('dc:title', 			$pPHPPowerPoint->getProperties()->getTitle());

			// dc:description
			$objWriter->writeElement('dc:description', 		$pPHPPowerPoint->getProperties()->getDescription());

			// dc:subject
			$objWriter->writeElement('dc:subject', 			$pPHPPowerPoint->getProperties()->getSubject());

			// cp:keywords
			$objWriter->writeElement('cp:keywords', 		$pPHPPowerPoint->getProperties()->getKeywords());

			// cp:category
			$objWriter->writeElement('cp:category', 		$pPHPPowerPoint->getProperties()->getCategory());

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}
}
