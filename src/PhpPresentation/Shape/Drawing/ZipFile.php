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

namespace PhpOffice\PhpPresentation\Shape\Drawing;

use PhpOffice\Common\File as CommonFile;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use ZipArchive;

class ZipFile extends AbstractDrawingAdapter
{
    /**
     * @var string
     */
    protected $path;

    /**
     * Get Path.
     */
    public function getPath(): string
    {
        return $this->path;
    }

    /**
     * Set Path.
     *
     * @param string $pValue File path
     */
    public function setPath(string $pValue = ''): self
    {
        $this->path = $pValue;

        return $this;
    }

    public function getContents(): string
    {
        if (!CommonFile::fileExists($this->getZipFileOut())) {
            throw new FileNotFoundException($this->getZipFileOut());
        }

        $imageZip = new ZipArchive();
        $imageZip->open($this->getZipFileOut());
        $imageContents = $imageZip->getFromName($this->getZipFileIn());
        $imageZip->close();
        unset($imageZip);

        return $imageContents;
    }

    public function getExtension(): string
    {
        return pathinfo($this->getZipFileIn(), PATHINFO_EXTENSION);
    }

    public function getMimeType(): string
    {
        if (!CommonFile::fileExists($this->getZipFileOut())) {
            throw new FileNotFoundException($this->getZipFileOut());
        }
        $oArchive = new ZipArchive();
        $oArchive->open($this->getZipFileOut());
        if (!function_exists('getimagesizefromstring')) {
            $uri = 'data://application/octet-stream;base64,' . base64_encode($oArchive->getFromName($this->getZipFileIn()));
            $image = getimagesize($uri);
        } else {
            $image = getimagesizefromstring($oArchive->getFromName($this->getZipFileIn()));
        }

        return image_type_to_mime_type($image[2]);
    }

    public function getIndexedFilename(): string
    {
        $output = pathinfo($this->getZipFileIn(), PATHINFO_FILENAME);
        $output = str_replace('.' . $this->getExtension(), '', $output);
        $output .= $this->getImageIndex();
        $output .= '.' . $this->getExtension();
        $output = str_replace(' ', '_', $output);

        return $output;
    }

    protected function getZipFileOut(): string
    {
        $path = str_replace('zip://', '', $this->getPath());
        $path = explode('#', $path);

        return empty($path[0]) ? '' : $path[0];
    }

    protected function getZipFileIn(): string
    {
        $path = str_replace('zip://', '', $this->getPath());
        $path = explode('#', $path);

        return empty($path[1]) ? '' : $path[1];
    }
}
