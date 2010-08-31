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
 * @package    PHPPowerPoint_Writer
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Writer_Serialized
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_Serialized implements PHPPowerPoint_Writer_IWriter
{
	/**
	 * Private PHPPowerPoint
	 *
	 * @var PHPPowerPoint
	 */
	private $_presentation;

    /**
     * Create a new PHPPowerPoint_Writer_Serialized
     *
	 * @param 	PHPPowerPoint	$pPHPPowerPoint
     */
    public function __construct(PHPPowerPoint $pPHPPowerPoint = null)
    {
    	// Assign PHPPowerPoint
		$this->setPHPPowerPoint($pPHPPowerPoint);
    }

	/**
	 * Save PHPPowerPoint to file
	 *
	 * @param 	string 		$pFileName
	 * @throws 	Exception
	 */
	public function save($pFilename = null)
	{
		if (!is_null($this->_presentation)) {
			// Create new ZIP file and open it for writing
			$objZip = new ZipArchive();

			// Try opening the ZIP file
			if ($objZip->open($pFilename, ZIPARCHIVE::OVERWRITE) !== true) {
				if ($objZip->open($pFilename, ZIPARCHIVE::CREATE) !== true) {
					throw new Exception("Could not open " . $pFilename . " for writing.");
				}
			}

			// Add media
			$slideCount = $this->_presentation->getSlideCount();
			for ($i = 0; $i < $slideCount; ++$i) {
				for ($j = 0; $j < $this->_presentation->getSlide($i)->getShapeCollection()->count(); ++$j) {
					if ($this->_presentation->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof PHPPowerPoint_Shape_BaseDrawing) {
						$imgTemp = $this->_presentation->getSlide($i)->getShapeCollection()->offsetGet($j);
						$objZip->addFromString('media/' . $imgTemp->getFilename(), file_get_contents($imgTemp->getPath()));
					}
				}
			}

			// Add PHPPowerPoint.xml to the document, which represents a PHP serialized PHPPowerPoint object
			$objZip->addFromString('PHPPowerPoint.xml', $this->_writeSerialized($this->_presentation, $pFilename));

			// Close file
			if ($objZip->close() === false) {
				throw new Exception("Could not close zip file $pFilename.");
			}
		} else {
			throw new Exception("PHPPowerPoint object unassigned.");
		}
	}

	/**
	 * Get PHPPowerPoint object
	 *
	 * @return PHPPowerPoint
	 * @throws Exception
	 */
	public function getPHPPowerPoint() {
		if (!is_null($this->_presentation)) {
			return $this->_presentation;
		} else {
			throw new Exception("No PHPPowerPoint assigned.");
		}
	}

	/**
	 * Get PHPPowerPoint object
	 *
	 * @param 	PHPPowerPoint 	$pPHPPowerPoint	PHPPowerPoint object
	 * @throws	Exception
	 * @return PHPPowerPoint_Writer_Serialized
	 */
	public function setPHPPowerPoint(PHPPowerPoint $pPHPPowerPoint = null) {
		$this->_presentation = $pPHPPowerPoint;
		return $this;
	}

	/**
	 * Serialize PHPPowerPoint object to XML
	 *
	 * @param 	PHPPowerPoint	$pPHPPowerPoint
	 * @param 	string		$pFilename
	 * @return 	string 		XML Output
	 * @throws 	Exception
	 */
	private function _writeSerialized(PHPPowerPoint $pPHPPowerPoint = null, $pFilename = '')
	{
		// Clone $pPHPPowerPoint
		$pPHPPowerPoint = clone $pPHPPowerPoint;

		// Update media links
		$slideCount = $pPHPPowerPoint->getSlideCount();
		for ($i = 0; $i < $slideCount; ++$i) {
			for ($j = 0; $j < $pPHPPowerPoint->getSlide($i)->getShapeCollection()->count(); ++$j) {
				if ($pPHPPowerPoint->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof PHPPowerPoint_Shape_BaseDrawing) {
					$imgTemp =& $pPHPPowerPoint->getSlide($i)->getShapeCollection()->offsetGet($j);
					$imgTemp->setPath('zip://' . $pFilename . '#media/' . $imgTemp->getFilename(), false);
				}
			}
		}

		// Create XML writer
		$objWriter = new xmlWriter();
		$objWriter->openMemory();
		$objWriter->setIndent(true);

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// PHPPowerPoint
		$objWriter->startElement('PHPPowerPoint');
		$objWriter->writeAttribute('version', '##VERSION##');

			// Comment
			$objWriter->writeComment('This file has been generated using PHPPowerPoint v##VERSION## (http://www.codeplex.com/PHPPowerPoint). It contains a base64 encoded serialized version of the PHPPowerPoint internal object.');

			// Data
			$objWriter->startElement('data');
				$objWriter->writeCData( base64_encode(serialize($pPHPPowerPoint)) );
			$objWriter->endElement();

		$objWriter->endElement();

		// Return
		return $objWriter->outputMemory(true);
	}
}
