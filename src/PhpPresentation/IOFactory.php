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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation;

use PhpOffice\PhpPresentation\Reader\ReaderInterface;
use PhpOffice\PhpPresentation\Writer\WriterInterface;
use ReflectionClass;

/**
 * IOFactory.
 */
class IOFactory
{
    /**
     * Autoresolve classes.
     *
     * @var array<int, string>
     */
    private static $autoResolveClasses = ['Serialized', 'ODPresentation', 'PowerPoint97', 'PowerPoint2007'];

    /**
     * Create writer.
     *
     * @throws \Exception
     */
    public static function createWriter(PhpPresentation $phpPresentation, string $name = 'PowerPoint2007'): WriterInterface
    {
        $class = 'PhpOffice\\PhpPresentation\\Writer\\' . $name;

        return self::loadClass($class, $name, 'writer', $phpPresentation);
    }

    /**
     * Create reader.
     *
     * @throws \Exception
     */
    public static function createReader(string $name = ''): ReaderInterface
    {
        $class = 'PhpOffice\\PhpPresentation\\Reader\\' . $name;

        return self::loadClass($class, $name, 'reader');
    }

    /**
     * Loads PhpPresentation from file using automatic \PhpOffice\PhpPresentation\Reader\ReaderInterface resolution.
     *
     * @throws \Exception
     */
    public static function load(string $pFilename): PhpPresentation
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
     * Load class.
     *
     * @return mixed
     *
     * @throws \ReflectionException
     */
    private static function loadClass(string $class, string $name, string $type, PhpPresentation $phpPresentation = null)
    {
        if (class_exists($class) && self::isConcreteClass($class)) {
            if (is_null($phpPresentation)) {
                return new $class();
            } else {
                return new $class($phpPresentation);
            }
        } else {
            throw new \Exception('"' . $name . '" is not a valid ' . $type . '.');
        }
    }

    /**
     * Is it a concrete class?
     *
     * @throws \ReflectionException
     */
    private static function isConcreteClass(string $class): bool
    {
        $reflection = new ReflectionClass($class);

        return !$reflection->isAbstract() && !$reflection->isInterface();
    }
}
