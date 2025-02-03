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

namespace PhpOffice\PhpPresentation\Slide\Background;

use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Slide\AbstractBackground;

class Image extends AbstractBackground
{
    /**
     * @var string
     */
    public $relationId;

    /**
     * Path.
     *
     * @var string
     */
    protected $path;

    /**
     * @var int
     */
    protected $height;

    /**
     * @var int
     */
    protected $width;

    /**
     * @var string
     */
    protected $extension;

    /**
     * Get Path.
     */
    public function getPath(): ?string
    {
        return $this->path;
    }

    /**
     * Set Path.
     *
     * @param string $pValue File path
     * @param bool $pVerifyFile Verify file
     *
     * @return self
     */
    public function setPath(string $pValue = '', bool $pVerifyFile = true)
    {
        if ($pVerifyFile) {
            if (!file_exists($pValue)) {
                throw new FileNotFoundException($pValue);
            }

            if (0 == $this->width && 0 == $this->height) {
                // Get width/height
                [$this->width, $this->height] = getimagesize($pValue);
            }
        }
        $this->path = $pValue;

        return $this;
    }

    /**
     * Set Extension.
     *
     * @param string $pValue File Extension
     */
    public function setExtension(string $pValue): self
    {
        $this->extension = $pValue;

        return $this;
    }

    /**
     * Get Filename.
     */
    public function getFilename(): string
    {
        return $this->path ? basename($this->path) : '';
    }

    /**
     * Get Extension.
     */
    public function getExtension(): string
    {
        if ($this->extension) {
            return $this->extension;
        }
        $exploded = explode('.', $this->getFilename());

        return $exploded[count($exploded) - 1];
    }

    /**
     * Get indexed filename (using image index).
     *
     * @param string $numSlide
     *
     * @return string
     */
    public function getIndexedFilename($numSlide)
    {
        return 'background_' . $numSlide . '.' . $this->getExtension();
    }
}
