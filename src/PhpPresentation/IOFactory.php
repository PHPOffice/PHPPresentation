<?php
/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation;

/**
 * IOFactory
 */
class IOFactory
{
    /**
     * Autoresolve classes
     *
     * @var array
     */
    private static $autoResolveClasses = array('Serialized', 'ODPresentation', 'PowerPoint97', 'PowerPoint2007');

    /**
     * Create writer
     *
     * @param PhpPresentation $phpPresentation
     * @param string $name
     * @return \PhpOffice\PhpPresentation\Writer\WriterInterface
     */
    public static function createWriter(PhpPresentation $phpPresentation, $name = 'PowerPoint2007')
    {
        $class = 'PhpOffice\\PhpPresentation\\Writer\\' . $name;
        return self::loadClass($class, $name, 'writer', $phpPresentation);
    }

    /**
     * Create reader
     *
     * @param  string $name
     * @return \PhpOffice\PhpPresentation\Reader\ReaderInterface
     */
    public static function createReader($name = '')
    {
        $class = 'PhpOffice\\PhpPresentation\\Reader\\' . $name;
        return self::loadClass($class, $name, 'reader');
    }

    /**
     * Loads PhpPresentation from file using automatic \PhpOffice\PhpPresentation\Reader\ReaderInterface resolution
     *
     * @param  string        $pFilename
     * @return PhpPresentation
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

        throw new \Exception("Could not automatically determine \PhpOffice\PhpPresentation\Reader\ReaderInterface for file.");
    }

    /**
     * Load class
     *
     * @param string $class
     * @param string $name
     * @param string $type
     * @param \PhpOffice\PhpPresentation\PhpPresentation $phpPresentation
     * @throws \Exception
     * @return
     */
    private static function loadClass($class, $name, $type, PhpPresentation $phpPresentation = null)
    {
        if (class_exists($class) && self::isConcreteClass($class)) {
            if (is_null($phpPresentation)) {
                return new $class();
            } else {
                return new $class($phpPresentation);
            }
        } else {
            throw new \Exception('"'.$name.'" is not a valid '.$type.'.');
        }
    }

    /**
     * Is it a concrete class?
     *
     * @param string $class
     * @return bool
     */
    private static function isConcreteClass($class)
    {
        $reflection = new \ReflectionClass($class);

        return !$reflection->isAbstract() && !$reflection->isInterface();
    }
}
