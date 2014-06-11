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
            throw new Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        return $this->fileSupportsUnserializePHPPowerPoint($pFilename);
    }

    /**
     * Loads PHPPowerPoint Serialized file
     *
     * @param  string        $pFilename
     * @return PHPPowerPoint
     * @throws Exception
     */
    public function load($pFilename)
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePHPPowerPoint($pFilename)) {
            throw new Exception("Invalid file format for PhpOffice\PhpPowerpoint\Reader\Serialized: " . $pFilename . ".");
        }

        return $this->_loadSerialized($pFilename);
    }

    /**
     * Load PHPPowerPoint Serialized file
     *
     * @param  string        $pFilename
     * @return PHPPowerPoint
     */
    private function _loadSerialized($pFilename)
    {
        $xmlData = simplexml_load_string(file_get_contents("zip://$pFilename#PHPPowerPoint.xml"));
        $excel   = unserialize(base64_decode((string) $xmlData->data));

        // Update media links
        for ($i = 0; $i < $excel->getSlideCount(); ++$i) {
            for ($j = 0; $j < $excel->getSlide($i)->getShapeCollection()->count(); ++$j) {
                if ($excel->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof BaseDrawing) {
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
     * @param  string    $pFilename
     * @throws Exception
     * @return boolean
     */
    public function fileSupportsUnserializePHPPowerPoint($pFilename = '')
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        // File exists, does it contain PHPPowerPoint.xml?
        return File::file_exists("zip://$pFilename#PHPPowerPoint.xml");
    }
}
