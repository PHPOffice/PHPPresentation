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

class File extends AbstractDrawingAdapter
{
    /**
     * @var string
     */
    protected $path = '';

    /**
     * @var string Name of the file
     */
    protected $fileName = '';

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
     * @param bool $pVerifyFile Verify file
     */
    public function setPath(string $pValue = '', bool $pVerifyFile = true): self
    {
        if ($pVerifyFile) {
            if (!file_exists($pValue)) {
                throw new FileNotFoundException($pValue);
            }
        }
        $this->path = $pValue;

        if ($pVerifyFile) {
            if (0 == $this->width && 0 == $this->height) {
                [$this->width, $this->height] = getimagesize($this->getPath());
            }
        }

        return $this;
    }

    public function getFileName(): string
    {
        return $this->fileName;
    }

    public function setFileName(string $fileName): self
    {
        $this->fileName = $fileName;

        return $this;
    }

    public function getContents(): string
    {
        return CommonFile::fileGetContents($this->getPath());
    }

    public function getExtension(): string
    {
        return pathinfo($this->getPath(), PATHINFO_EXTENSION);
    }

    public function getMimeType(): string
    {
        if (!CommonFile::fileExists($this->getPath())) {
            throw new FileNotFoundException($this->getPath());
        }
        $image = getimagesizefromstring(CommonFile::fileGetContents($this->getPath()));

        if (is_array($image)) {
            return image_type_to_mime_type($image[2]);
        }

        return mime_content_type($this->getPath());
    }

    public function getIndexedFilename(): string
    {
        $output = str_replace('.' . $this->getExtension(), '', pathinfo($this->getPath(), PATHINFO_FILENAME));
        $output .= $this->getImageIndex();
        $output .= '.' . $this->getExtension();
        $output = str_replace(' ', '_', $output);

        return $output;
    }
}
