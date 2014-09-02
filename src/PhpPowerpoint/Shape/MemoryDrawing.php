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

namespace PhpOffice\PhpPowerpoint\Shape;

use PhpOffice\PhpPowerpoint\ComparableInterface;

/**
 * Memory drawing shape
 */
class MemoryDrawing extends AbstractDrawing implements ComparableInterface
{
    /* Rendering functions */
    const RENDERING_DEFAULT = 'imagepng';
    const RENDERING_PNG = 'imagepng';
    const RENDERING_GIF = 'imagegif';
    const RENDERING_JPEG = 'imagejpeg';

    /* MIME types */
    const MIMETYPE_DEFAULT = 'image/png';
    const MIMETYPE_PNG = 'image/png';
    const MIMETYPE_GIF = 'image/gif';
    const MIMETYPE_JPEG = 'image/jpeg';

    /**
     * Image resource
     *
     * @var resource
     */
    private $imageResource;

    /**
     * Rendering function
     *
     * @var string
     */
    private $renderingFunction;

    /**
     * Mime type
     *
     * @var string
     */
    private $mimeType;

    /**
     * Unique name
     *
     * @var string
     */
    private $uniqueName;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Slide\MemoryDrawing
     */
    public function __construct()
    {
        // Initialise values
        $this->setImageResource(null);
        $this->setRenderingFunction(self::RENDERING_DEFAULT);
        $this->setMimeType(self::MIMETYPE_DEFAULT);
        $this->uniqueName = md5(rand(0, 9999) . time() . rand(0, 9999));

        // Initialize parent
        parent::__construct();
    }

    /**
     * Get image resource
     *
     * @return resource
     */
    public function getImageResource()
    {
        return $this->imageResource;
    }

    /**
     * Set image resource
     *
     * @param $value resource
     * @return \PhpOffice\PhpPowerpoint\Shape\MemoryDrawing
     */
    public function setImageResource($value = null)
    {
        $this->imageResource = $value;

        if (!is_null($this->imageResource)) {
            // Get width/height
            $this->width  = imagesx($this->imageResource);
            $this->height = imagesy($this->imageResource);
        }

        return $this;
    }

    /**
     * Get rendering function
     *
     * @return string
     */
    public function getRenderingFunction()
    {
        return $this->renderingFunction;
    }

    /**
     * Set rendering function
     *
     * @param  string                            $value
     * @return \PhpOffice\PhpPowerpoint\Shape\MemoryDrawing
     */
    public function setRenderingFunction($value = self::RENDERING_DEFAULT)
    {
        $this->renderingFunction = $value;

        return $this;
    }

    /**
     * Get mime type
     *
     * @return string
     */
    public function getMimeType()
    {
        return $this->mimeType;
    }

    /**
     * Set mime type
     *
     * @param  string                            $value
     * @return \PhpOffice\PhpPowerpoint\Shape\MemoryDrawing
     */
    public function setMimeType($value = self::MIMETYPE_DEFAULT)
    {
        $this->mimeType = $value;

        return $this;
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename()
    {
        return $this->uniqueName . $this->getImageIndex() . '.' . $this->getExtension();
    }

    /**
     * Get extension
     */
    public function getExtension()
    {
        $extension = strtolower($this->getMimeType());
        $extension = explode('/', $extension);
        return $extension[1];
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->renderingFunction . $this->mimeType . $this->uniqueName . parent::getHashCode() . __CLASS__);
    }
}
