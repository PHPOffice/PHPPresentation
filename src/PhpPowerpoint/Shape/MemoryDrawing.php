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

use PhpOffice\PhpPowerpoint\Shape\BaseDrawing;
use PhpOffice\PhpPowerpoint\IComparable;

/**
 * PHPPowerPoint_Shape_MemoryDrawing
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class MemoryDrawing extends BaseDrawing implements IComparable
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
    private $_imageResource;

    /**
     * Rendering function
     *
     * @var string
     */
    private $_renderingFunction;

    /**
     * Mime type
     *
     * @var string
     */
    private $_mimeType;

    /**
     * Unique name
     *
     * @var string
     */
    private $_uniqueName;

    /**
     * Create a new PHPPowerPoint_Slide_MemoryDrawing
     */
    public function __construct()
    {
        // Initialise values
        $this->_imageResource     = null;
        $this->_renderingFunction = self::RENDERING_DEFAULT;
        $this->_mimeType          = self::MIMETYPE_DEFAULT;
        $this->_uniqueName        = md5(rand(0, 9999) . time() . rand(0, 9999));

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
        return $this->_imageResource;
    }

    /**
     * Set image resource
     *
     * @param                                    $value resource
     * @return PHPPowerPoint_Shape_MemoryDrawing
     */
    public function setImageResource($value = null)
    {
        $this->_imageResource = $value;

        if (!is_null($this->_imageResource)) {
            // Get width/height
            $this->_width  = imagesx($this->_imageResource);
            $this->_height = imagesy($this->_imageResource);
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
        return $this->_renderingFunction;
    }

    /**
     * Set rendering function
     *
     * @param  string                            $value
     * @return PHPPowerPoint_Shape_MemoryDrawing
     */
    public function setRenderingFunction($value = self::RENDERING_DEFAULT)
    {
        $this->_renderingFunction = $value;

        return $this;
    }

    /**
     * Get mime type
     *
     * @return string
     */
    public function getMimeType()
    {
        return $this->_mimeType;
    }

    /**
     * Set mime type
     *
     * @param  string                            $value
     * @return PHPPowerPoint_Shape_MemoryDrawing
     */
    public function setMimeType($value = self::MIMETYPE_DEFAULT)
    {
        $this->_mimeType = $value;

        return $this;
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename()
    {
        $extension = strtolower($this->getMimeType());
        $extension = explode('/', $extension);
        $extension = $extension[1];

        return $this->_uniqueName . $this->getImageIndex() . '.' . $extension;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->_renderingFunction . $this->_mimeType . $this->_uniqueName . parent::getHashCode() . __CLASS__);
    }

    /**
     * Hash index
     *
     * @var string
     */
    private $_hashIndex;

    /**
     * Get hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return string Hash index
     */
    public function getHashIndex()
    {
        return $this->_hashIndex;
    }

    /**
     * Set hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param string $value Hash index
     */
    public function setHashIndex($value)
    {
        $this->_hashIndex = $value;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
}
