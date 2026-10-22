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
     * @var resource
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
     * The bytes the picture was read from.
     *
     * A picture that was made rather than read has none, and its contents are rendered
     * from its resource instead.
     *
     * @var string
     */
    protected $contents = '';

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
     * @return resource
     */
    public function getImageResource()
    {
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
        // the bytes that were read describe the resource they were decoded into, so a
        // resource that is not that one has no bytes
        $this->contents = '';

        if (null !== $this->imageResource && false !== $this->imageResource) {
            // Get width/height
            $this->width = imagesx($this->imageResource);
            $this->height = imagesy($this->imageResource);
        }

        return $this;
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

    /**
     * Set the bytes the picture was read from.
     *
     * The bytes are the ones the file holds, so that a picture read out of a package
     * is given back as it was read rather than as an encoding of it made again.
     */
    public function setContents(string $contents): self
    {
        $this->contents = $contents;

        return $this;
    }

    public function getContents(): string
    {
        if ('' !== $this->contents) {
            return $this->contents;
        }

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
        if ('' !== $this->path) {
            return $this->nameFromPath($this->getExtension());
        }

        return $this->uniqueName . $this->getImageIndex() . '.' . $this->getExtension();
    }

    /**
     * @var string
     */
    protected $path = '';

    /**
     * Get Path.
     */
    public function getPath(): string
    {
        return $this->path;
    }

    public function setPath(string $path): self
    {
        $this->path = $path;

        return $this;
    }
}
