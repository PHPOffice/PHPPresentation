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
     * @var bool Flag indicating if this is a temporary file that should be cleaned up
     */
    protected $isTemporaryFile = false;

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

    /**
     * Set whether this is a temporary file that should be cleaned up
     *
     * @param bool $isTemporary
     * @return self
     */
    public function setIsTemporaryFile(bool $isTemporary): self
    {
        $this->isTemporaryFile = $isTemporary;
        return $this;
    }

    /**
     * Check if this is a temporary file that should be cleaned up
     *
     * @return bool
     */
    public function isTemporaryFile(): bool
    {
        return $this->isTemporaryFile;
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

    /**
     * {@inheritDoc}
     */
    public function loadFromContent(string $content, string $fileName = '', string $prefix = 'PhpPresentation'): AbstractDrawingAdapter
    {
        // Create temporary file
        $tmpFile = tempnam(sys_get_temp_dir(), $prefix);
        file_put_contents($tmpFile, $content);

        // Set path and mark as temporary
        $this->setPath($tmpFile);
        $this->setIsTemporaryFile(true);

        // Set filename if provided
        if (!empty($fileName)) {
            $this->setFileName($fileName);
        }

        return $this;
    }

    /**
     * Clean up resources when object is destroyed
     */
    public function __destruct()
    {
        // Remove temporary file if needed
        if ($this->isTemporaryFile() && $this->path && file_exists($this->path)) {
            @unlink($this->path);
        }
    }
}
