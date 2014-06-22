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

namespace PhpOffice\PhpPowerpoint\Reader;

use PhpOffice\PhpPowerpoint\Reader\IReader;
use PhpOffice\PhpPowerpoint\Shared\File;
use PhpOffice\PhpPowerpoint\Shape\BaseDrawing;

/**
 * Serialized format reader
 */
class Serialized implements IReader
{
    /**
     * Can the current PHPPowerPoint_Reader_IReader read the file?
     *
     * @param  string  $pFilename
     * @return boolean
     */
    public function canRead($pFilename)
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        return $this->fileSupportsUnserializePHPPowerPoint($pFilename);
    }

    /**
     * Loads PHPPowerPoint Serialized file
     *
     * @param  string        $pFilename
     * @return PHPPowerPoint
     * @throws \Exception
     */
    public function load($pFilename)
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePHPPowerPoint($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPowerpoint\Reader\Serialized: " . $pFilename . ".");
        }

        return $this->loadSerialized($pFilename);
    }

    /**
     * Does a file support UnserializePHPPowerPoint ?
     *
     * @param  string    $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function fileSupportsUnserializePHPPowerPoint($pFilename = '')
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        // File exists, does it contain PHPPowerPoint.xml?
        return File::fileExists("zip://$pFilename#PHPPowerPoint.xml");
    }

    /**
     * Load PHPPowerPoint Serialized file
     *
     * @param  string        $pFilename
     * @return PHPPowerPoint
     */
    private function loadSerialized($pFilename)
    {
        $oArchive = new \ZipArchive();
        if ($oArchive->open($pFilename) === true) {
            $xmlContent = $oArchive->getFromName('PHPPowerPoint.xml');

            if (!empty($xmlContent)) {
                $xmlData = simplexml_load_string($xmlContent);
                $file    = unserialize(base64_decode((string) $xmlData->data));

                // Update media links
                for ($i = 0; $i < $file->getSlideCount(); ++$i) {
                    for ($j = 0; $j < $file->getSlide($i)->getShapeCollection()->count(); ++$j) {
                        if ($file->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof BaseDrawing) {
                            $file->getSlide($i)->getShapeCollection()->offsetGet($j)->setPath('zip://' . $pFilename . '#media/' . $file->getSlide($i)->getShapeCollection()->offsetGet($j)->getFilename(), false);
                        }
                    }
                }

                $oArchive->close();
                return $file;
            }
        }
    }
}
