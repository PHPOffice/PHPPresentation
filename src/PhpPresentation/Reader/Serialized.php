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

namespace PhpOffice\PhpPresentation\Reader;

use Exception;
use PhpOffice\Common\File;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Drawing\File as DrawingFile;
use ZipArchive;

/**
 * Serialized format reader.
 */
class Serialized implements ReaderInterface
{
    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     *
     * @throws \Exception
     */
    public function canRead(string $pFilename): bool
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     *
     * @throws \Exception
     */
    public function fileSupportsUnserializePhpPresentation(string $pFilename = ''): bool
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception('Could not open ' . $pFilename . ' for reading! File does not exist.');
        }

        // File exists, does it contain PhpPresentation.xml?
        return File::fileExists("zip://$pFilename#PhpPresentation.xml");
    }

    /**
     * Loads PhpPresentation Serialized file.
     *
     * @throws \Exception
     */
    public function load(string $pFilename): PhpPresentation
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception('Could not open ' . $pFilename . ' for reading! File does not exist.');
        }

        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPresentation\Reader\Serialized: " . $pFilename . '.');
        }

        return $this->loadSerialized($pFilename);
    }

    /**
     * Load PhpPresentation Serialized file.
     */
    private function loadSerialized(string $pFilename): PhpPresentation
    {
        $oArchive = new ZipArchive();
        if (true !== $oArchive->open($pFilename)) {
            throw new Exception('');
        }

        $xmlContent = $oArchive->getFromName('PhpPresentation.xml');
        if (empty($xmlContent)) {
            throw new Exception(sprintf('The file %s in the serialized file %s is malformed', 'PhpPresentation.xml', $pFilename));
        }

        $xmlData = simplexml_load_string($xmlContent);
        $file = unserialize(base64_decode((string) $xmlData->data));

        // Update media links
        for ($i = 0; $i < $file->getSlideCount(); ++$i) {
            for ($j = 0; $j < $file->getSlide($i)->getShapeCollection()->count(); ++$j) {
                if ($file->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof AbstractDrawingAdapter) {
                    $imgTemp = $file->getSlide($i)->getShapeCollection()->offsetGet($j);
                    $imgPath = 'zip://' . $pFilename . '#media/' . $imgTemp->getImageIndex() . '/' . pathinfo($imgTemp->getPath(), PATHINFO_BASENAME);
                    if ($imgTemp instanceof DrawingFile) {
                        $imgTemp->setPath($imgPath, false);
                    } else {
                        $imgTemp->setPath($imgPath);
                    }
                }
            }
        }

        $oArchive->close();

        return $file;
    }
}
