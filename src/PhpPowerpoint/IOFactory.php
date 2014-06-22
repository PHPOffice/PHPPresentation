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

namespace PhpOffice\PhpPowerpoint;

/**
 * PHPPowerPoint_IOFactory
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class IOFactory
{
    /**
     * Autoresolve classes
     *
     * @var array
     */
    private static $autoResolveClasses = array('Serialized');

    /**
     * Is it a concrete class?
     *
     * @param string $class
     */
    private static function isConcreteClass($class)
    {
        $reflection = new \ReflectionClass($class);

        return !$reflection->isAbstract() && !$reflection->isInterface();
    }

    /**
     * Load class
     *
     * @param string $class
     * @param string $name
     * @param string $type
     * @param \PhpOffice\PhpPowerpoint\PhpPowerpoint $phpPowerPoint
     */
    private static function loadClass($class, $name, $type, PhpPowerpoint $phpPowerPoint = null)
    {
        if (class_exists($class) && self::isConcreteClass($class)) {
            if (is_null($phpPowerPoint)) {
                return new $class();
            } else {
                return new $class($phpPowerPoint);
            }
        } else {
            throw new \Exception('"'.$name.'" is not a valid '.$type.'.');
        }
    }

    /**
     * Create writer
     *
     * @param PHPPowerPoint $PHPPowerPoint
     * @param string $name
     * @return \PhpOffice\PhpPowerpoint\PhpPowerpoint\Writer\IWriter
     */
    public static function createWriter(PhpPowerpoint $phpPowerPoint, $name = 'PowerPoint2007')
    {
        $class = 'PhpOffice\\PhpPowerpoint\\Writer\\' . $name;
        return self::loadClass($class, $name, 'writer', $phpPowerPoint);
    }

    /**
     * Create reader
     *
     * @param  string $name
     * @return PHPPowerPoint_Reader_IReader
     */
    public static function createReader($name = '')
    {
        $class = 'PhpOffice\\PhpPowerpoint\\Reader\\' . $name;
        return self::loadClass($class, $name, 'reader');
    }

    /**
     * Loads PHPPowerPoint from file using automatic PHPPowerPoint_Reader_IReader resolution
     *
     * @param  string        $pFilename
     * @return PHPPowerPoint
     * @throws \Exception
     */
    public static function load($pFilename)
    {
        // Try loading using self::$autoResolveClasses
        foreach (self::$autoResolveClasses as $autoResolveClass) {
            $reader = self::createReader($autoResolveClass);
            if ($reader->canRead($pFilename)) {
                return $reader->load($pFilename);
            }
        }

        throw new \Exception("Could not automatically determine PHPPowerPoint_Reader_IReader for file.");
    }
}
