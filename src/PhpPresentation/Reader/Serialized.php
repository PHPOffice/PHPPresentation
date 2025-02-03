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
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Reader;

use PhpOffice\Common\File;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
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
     */
    public function canRead(string $pFilename): bool
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     */
    public function fileSupportsUnserializePhpPresentation(string $pFilename): bool
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new FileNotFoundException($pFilename);
        }

        // File exists, does it contain PhpPresentation.xml?
        return File::fileExists("zip://$pFilename#PhpPresentation.xml");
    }

    /**
     * Loads PhpPresentation Serialized file.
     */
    public function load(string $pFilename, int $flags = 0): PhpPresentation
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new FileNotFoundException($pFilename);
        }

        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new InvalidFileFormatException($pFilename, self::class);
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
            throw new InvalidFileFormatException($pFilename, self::class);
        }

        $xmlContent = $oArchive->getFromName('PhpPresentation.xml');
        if (empty($xmlContent)) {
            throw new InvalidFileFormatException($pFilename, self::class, 'The file PhpPresentation.xml is malformed');
        }

        $xmlData = simplexml_load_string($xmlContent);
        $file = unserialize(base64_decode((string) $xmlData->data));

        // Update media links
        for ($i = 0; $i < $file->getSlideCount(); ++$i) {
            foreach ($file->getSlide($i)->getShapeCollection() as $shape) {
                if ($shape instanceof AbstractDrawingAdapter) {
                    $imgPath = 'zip://' . $pFilename . '#media/' . $shape->getImageIndex() . '/' . pathinfo($shape->getPath(), PATHINFO_BASENAME);
                    if ($shape instanceof DrawingFile) {
                        $shape->setPath($imgPath, false);
                    } else {
                        $shape->setPath($imgPath);
                    }
                }
            }
        }

        $oArchive->close();

        return $file;
    }
}
