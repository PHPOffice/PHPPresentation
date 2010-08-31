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
 * @package    PHPPowerPoint_Reader
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/** PHPPowerPoint root directory */
if (!defined('PHPPOWERPOINT_ROOT')) {
	/**
	 * @ignore
	 */
	define('PHPPOWERPOINT_ROOT', dirname(__FILE__) . '/../../');
	require(PHPPOWERPOINT_ROOT . 'PHPPowerPoint/Autoloader.php');
	PHPPowerPoint_Autoloader::Register();
	PHPPowerPoint_Shared_ZipStreamWrapper::register();
}


/**
 * PHPPowerPoint_Reader_Serialized
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Reader
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Reader_Serialized implements PHPPowerPoint_Reader_IReader
{
	/**
	 * Can the current PHPPowerPoint_Reader_IReader read the file?
	 *
	 * @param 	string 		$pFileName
	 * @return 	boolean
	 */
	public function canRead($pFilename)
	{
		// Check if file exists
		if (!file_exists($pFilename)) {
			throw new Exception("Could not open " . $pFilename . " for reading! File does not exist.");
		}

		return $this->fileSupportsUnserializePHPPowerPoint($pFilename);
	}

	/**
	 * Loads PHPPowerPoint Serialized file
	 *
	 * @param 	string 		$pFilename
	 * @return 	PHPPowerPoint
	 * @throws 	Exception
	 */
	public function load($pFilename)
	{
		// Check if file exists
		if (!file_exists($pFilename)) {
			throw new Exception("Could not open " . $pFilename . " for reading! File does not exist.");
		}

		// Unserialize... First make sure the file supports it!
		if (!$this->fileSupportsUnserializePHPPowerPoint($pFilename)) {
			throw new Exception("Invalid file format for PHPPowerPoint_Reader_Serialized: " . $pFilename . ".");
		}

		return $this->_loadSerialized($pFilename);
	}

	/**
	 * Load PHPPowerPoint Serialized file
	 *
	 * @param 	string 		$pFilename
	 * @return 	PHPPowerPoint
	 */
	private function _loadSerialized($pFilename) {
		$xmlData = simplexml_load_string(file_get_contents("zip://$pFilename#PHPPowerPoint.xml"));
		$excel = unserialize(base64_decode((string)$xmlData->data));

		// Update media links
		for ($i = 0; $i < $excel->getSlideCount(); ++$i) {
			for ($j = 0; $j < $excel->getSlide($i)->getShapeCollection()->count(); ++$j) {
				if ($excel->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof PHPExcl_Shape_BaseDrawing) {
					$imgTemp =& $excel->getSlide($i)->getShapeCollection()->offsetGet($j);
					$imgTemp->setPath('zip://' . $pFilename . '#media/' . $imgTemp->getFilename(), false);
				}
			}
		}

		return $excel;
	}

	/**
	 * Does a file support UnserializePHPPowerPoint ?
	 *
	 * @param 	string 		$pFilename
	 * @throws 	Exception
	 * @return 	boolean
	 */
	public function fileSupportsUnserializePHPPowerPoint($pFilename = '') {
		// Check if file exists
		if (!file_exists($pFilename)) {
			throw new Exception("Could not open " . $pFilename . " for reading! File does not exist.");
		}

		// File exists, does it contain PHPPowerPoint.xml?
		return PHPPowerPoint_Shared_File::file_exists("zip://$pFilename#PHPPowerPoint.xml");
	}
}
