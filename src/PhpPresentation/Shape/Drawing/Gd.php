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

use GdImage;

class Gd extends AbstractDrawingAdapter
{
    // Rendering functions
    public const RENDERING_DEFAULT = 'imagepng';
    public const RENDERING_PNG = 'imagepng';
    public const RENDERING_GIF = 'imagegif';
    public const RENDERING_JPEG = 'imagejpeg';

    // MIME types
    public const MIMETYPE_DEFAULT = 'image/png';
    public const MIMETYPE_PNG = 'image/png';
    public const MIMETYPE_GIF = 'image/gif';
    public const MIMETYPE_JPEG = 'image/jpeg';

    /**
     * Image resource.
     *
     * @var GdImage|resource|null
     */
    protected $imageResource;

    /**
     * Rendering function.
     *
     * @var string
     */
    protected $renderingFunction = self::RENDERING_DEFAULT;

    /**
     * Mime type.
     *
     * @var string
     */
    protected $mimeType = self::MIMETYPE_DEFAULT;

    /**
     * Unique name.
     *
     * @var string
     */
    protected $uniqueName;

    /**
     * @var bool Flag indicating if this is a temporary file that should be cleaned up
     */
    protected $isTemporaryFile = false;

    /**
     * Gd constructor.
     */
    public function __construct()
    {
        parent::__construct();
        $this->uniqueName = md5(mt_rand(0, 9999) . time() . mt_rand(0, 9999));
    }

    /**
     * Get image resource.
     *
     * @param bool $isTransient Avoid the image resource being stored in memory to avoid OOM
     * @return ?GdImage|?resource
     */
    public function getImageResource(bool $isTransient = false)
    {
        // Lazy load image resource if not already loaded
        if (!$this->imageResource) {
            $imageString = file_get_contents($this->getPath());
            if ($imageString === false) {
                return null; // Failed to read file
            }

            $image = imagecreatefromstring($imageString);
            if ($image === false) {
                return null; // Failed to create image resource
            }

            $this->setImageResource($image);
        }

        if ($isTransient) {
            // Create a new image resource and copy the original
            $width = imagesx($this->imageResource);
            $height = imagesy($this->imageResource);
            $imageCopy = imagecreatetruecolor($width, $height);

            // Preserve transparency for PNG/GIF images
            if (imageistruecolor($this->imageResource)) {
                imagealphablending($imageCopy, false);
                imagesavealpha($imageCopy, true);
            }

            // Copy the image data
            imagecopy($imageCopy, $this->imageResource, 0, 0, 0, 0, $width, $height);

            // Destroy the original resource to free memory
            imagedestroy($this->imageResource);
            $this->imageResource = null;

            return $imageCopy;
        }

        return $this->imageResource;
    }

    /**
     * Set image resource.
     *
     * @param null|false|GdImage|resource $value
     *
     * @return $this
     */
    public function setImageResource($value = null)
    {
        $this->imageResource = $value;
        if (!$this->imageResource) {
            return $this;
        }

        $this->getDimensions();

        return $this;
    }

    public function getDimensions(): array
    {
        // Lazy load dimensions
        if (!$this->width) {
            $this->width = imagesx($this->imageResource);
        }
        if (!$this->height) {
            $this->height = imagesy($this->imageResource);
        }

        return [$this->width, $this->height];
    }

    /**
     * Get rendering function.
     *
     * @return string
     */
    public function getRenderingFunction()
    {
        return $this->renderingFunction;
    }

    /**
     * Set rendering function.
     *
     * @param string $value
     *
     * @return $this
     */
    public function setRenderingFunction($value = self::RENDERING_DEFAULT)
    {
        $this->renderingFunction = $value;

        return $this;
    }

    /**
     * Get mime type.
     */
    public function getMimeType(): string
    {
        return $this->mimeType;
    }

    /**
     * Set mime type.
     *
     * @param string $value
     *
     * @return $this
     */
    public function setMimeType($value = self::MIMETYPE_DEFAULT)
    {
        $this->mimeType = $value;

        return $this;
    }

    public function getContents(): string
    {
        ob_start();
        if (self::MIMETYPE_DEFAULT === $this->getMimeType()) {
            imagealphablending($this->getImageResource(), false);
            imagesavealpha($this->getImageResource(), true);
        }
        call_user_func($this->getRenderingFunction(), $this->getImageResource());
        $imageContents = ob_get_contents();
        ob_end_clean();

        return $imageContents;
    }

    public function getExtension(): string
    {
        $extension = strtolower($this->getMimeType());
        $extension = explode('/', $extension);
        $extension = $extension[1];

        return $extension;
    }

    public function getIndexedFilename(): string
    {
        return $this->uniqueName . $this->getImageIndex() . '.' . $this->getExtension();
    }

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
     * @param string $path File path
     * @return self
     */
    public function setPath(string $path): self
    {
        $this->path = $path;
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

    /**
     * Clean up resources when object is destroyed
     */
    public function __destruct()
    {
        // Free GD image resource if it exists
        if ($this->imageResource) {
            imagedestroy($this->imageResource);
            $this->imageResource = null;
        }

        // Remove temporary file if needed
        if ($this->isTemporaryFile && !empty($this->path) && file_exists($this->path)) {
            @unlink($this->path);
        }
    }

    /**
     * {@inheritDoc}
     */
    public function loadFromContent(string $content, string $fileName = '', string $prefix = 'PhpPresentationGd'): AbstractDrawingAdapter
    {
        // Check if the content is a valid image
        $image = @imagecreatefromstring($content);
        if ($image === false) {
            return $this;
        }
        // Clean up the image resource to avoid memory leaks
        @imagedestroy($image);

        $tmpFile = tempnam(sys_get_temp_dir(), $prefix);
        file_put_contents($tmpFile, $content);

        // Set path and mark as temporary for automatic cleanup
        $this->setPath($tmpFile);
        $this->setIsTemporaryFile(true);

        if (!empty($fileName)) {
            $this->setName($fileName);
        }

        $info = getimagesizefromstring($content);
        if (isset($info['mime'])) {
            $this->setMimeType($info['mime']);
            $this->setRenderingFunction(str_replace('/', '', $info['mime']));
        }

        return $this;
    }
}
