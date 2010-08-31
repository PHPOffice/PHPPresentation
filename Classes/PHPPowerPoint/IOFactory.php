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
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/** PHPPowerPoint root directory */
if (!defined('PHPPOWERPOINT_ROOT')) {
	/**
	 * @ignore
	 */
	define('PHPPOWERPOINT_ROOT', dirname(__FILE__) . '/../');
	require(PHPPOWERPOINT_ROOT . 'PHPPowerPoint/Autoloader.php');
	PHPPowerPoint_Autoloader::Register();
	PHPPowerPoint_Shared_ZipStreamWrapper::register();
}


/**
 * PHPPowerPoint_IOFactory
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_IOFactory
{
	/**
	 * Search locations
	 *
	 * @var array
	 */
	private static $_searchLocations = array(
		array( 'type' => 'IWriter', 'path' => 'PHPPowerPoint/Writer/{0}.php', 'class' => 'PHPPowerPoint_Writer_{0}' ),
		array( 'type' => 'IReader', 'path' => 'PHPPowerPoint/Reader/{0}.php', 'class' => 'PHPPowerPoint_Reader_{0}' )
	);

	/**
	 * Autoresolve classes
	 *
	 * @var array
	 */
	private static $_autoResolveClasses = array(
		'Serialized'
	);

    /**
     * Private constructor for PHPPowerPoint_IOFactory
     */
    private function __construct() { }

    /**
     * Get search locations
     *
     * @return array
     */
	public static function getSearchLocations() {
		return self::$_searchLocations;
	}

	/**
	 * Set search locations
	 *
	 * @param array $value
	 * @throws Exception
	 */
	public static function setSearchLocations($value) {
		if (is_array($value)) {
			self::$_searchLocations = $value;
		} else {
			throw new Exception('Invalid parameter passed.');
		}
	}

	/**
	 * Add search location
	 *
	 * @param string $type			Example: IWriter
	 * @param string $location		Example: PHPPowerPoint/Writer/{0}.php
	 * @param string $classname 	Example: PHPPowerPoint_Writer_{0}
	 */
	public static function addSearchLocation($type = '', $location = '', $classname = '') {
		self::$_searchLocations[] = array( 'type' => $type, 'path' => $location, 'class' => $classname );
	}

	/**
	 * Create PHPPowerPoint_Writer_IWriter
	 *
	 * @param PHPPowerPoint $PHPPowerPoint
	 * @param string  $writerType	Example: PowerPoint2007
	 * @return PHPPowerPoint_Writer_IWriter
	 */
	public static function createWriter(PHPPowerPoint $PHPPowerPoint, $writerType = '') {
		// Search type
		$searchType = 'IWriter';

		// Include class
		foreach (self::$_searchLocations as $searchLocation) {
			if ($searchLocation['type'] == $searchType) {
				$className = str_replace('{0}', $writerType, $searchLocation['class']);
				$classFile = str_replace('{0}', $writerType, $searchLocation['path']);

				if (!class_exists($className)) {
					require_once($classFile);
				}

				$instance = new $className($PHPPowerPoint);
				if (!is_null($instance)) {
					return $instance;
				}
			}
		}

		// Nothing found...
		throw new Exception("No $searchType found for type $writerType");
	}

	/**
	 * Create PHPPowerPoint_Reader_IReader
	 *
	 * @param string $readerType	Example: PowerPoint2007
	 * @return PHPPowerPoint_Reader_IReader
	 */
	public static function createReader($readerType = '') {
		// Search type
		$searchType = 'IReader';

		// Include class
		foreach (self::$_searchLocations as $searchLocation) {
			if ($searchLocation['type'] == $searchType) {
				$className = str_replace('{0}', $readerType, $searchLocation['class']);
				$classFile = str_replace('{0}', $readerType, $searchLocation['path']);

				if (!class_exists($className)) {
					require_once($classFile);
				}

				$instance = new $className();
				if (!is_null($instance)) {
					return $instance;
				}
			}
		}

		// Nothing found...
		throw new Exception("No $searchType found for type $readerType");
	}

	/**
	 * Loads PHPPowerPoint from file using automatic PHPPowerPoint_Reader_IReader resolution
	 *
	 * @param 	string 		$pFileName
	 * @return	PHPPowerPoint
	 * @throws 	Exception
	 */
	public static function load($pFilename) {
		// Try loading using self::$_autoResolveClasses
		foreach (self::$_autoResolveClasses as $autoResolveClass) {
			$reader = self::createReader($autoResolveClass);
			if ($reader->canRead($pFilename)) {
				return $reader->load($pFilename);
			}
		}

		throw new Exception("Could not automatically determine PHPPowerPoint_Reader_IReader for file.");
	}
}
