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

namespace PhpOffice\PhpPresentation\Reader;

use PhpOffice\Common\File;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;

/**
 * Serialized format reader
 */
class Serialized implements ReaderInterface
{
    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     *
     * @param  string $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function canRead($pFilename)
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     *
     * @param  string    $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function fileSupportsUnserializePhpPresentation($pFilename = '')
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        // File exists, does it contain PhpPresentation.xml?
        return File::fileExists("zip://$pFilename#PhpPresentation.xml");
    }

    /**
     * Loads PhpPresentation Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     * @throws \Exception
     */
    public function load($pFilename)
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPresentation\Reader\Serialized: " . $pFilename . ".");
        }

        return $this->loadSerialized($pFilename);
    }

    /**
     * Load PhpPresentation Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     */
    private function loadSerialized($pFilename)
    {
        $oArchive = new \ZipArchive();
        if ($oArchive->open($pFilename) === true) {
            $xmlContent = $oArchive->getFromName('PhpPresentation.xml');

            if (!empty($xmlContent)) {
                $xmlData = simplexml_load_string($xmlContent);
                $file    = unserialize(base64_decode((string) $xmlData->data));

                // Update media links
                for ($i = 0; $i < $file->getSlideCount(); ++$i) {
                    for ($j = 0; $j < $file->getSlide($i)->getShapeCollection()->count(); ++$j) {
                        if ($file->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof AbstractDrawingAdapter) {
                            $file->getSlide($i)->getShapeCollection()->offsetGet($j)->setPath('zip://' . $pFilename . '#media/' . $file->getSlide($i)->getShapeCollection()->offsetGet($j)->getIndexedFilename(), false);
                        }
                    }
                }

                $oArchive->close();
                return $file;
            }
        }

        return null;
    }
}
