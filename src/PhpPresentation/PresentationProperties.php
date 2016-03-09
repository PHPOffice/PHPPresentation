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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */
namespace PhpOffice\PhpPresentation;

/**
 * \PhpOffice\PhpPresentation\PresentationProperties
 */
class PresentationProperties
{
    /*
     * @var boolean
     */
    protected $isLoopUntilEsc = false;

    /**
     * Mark as final
     * @var bool
     */
    protected $markAsFinal = false;

    /*
     * @var string
     */
    protected $thumbnail;

    /**
     * Zoom
     * @var float
     */
    protected $zoom = 1;
    
    /**
     * @return bool
     */
    public function isLoopContinuouslyUntilEsc()
    {
        return $this->isLoopUntilEsc;
    }
    
    /**
     * @param bool $value
     * @return \PhpOffice\PhpPresentation\PresentationProperties
     */
    public function setLoopContinuouslyUntilEsc($value = false)
    {
        if (is_bool($value)) {
            $this->isLoopUntilEsc = $value;
        }
        return $this;
    }
    
    /**
     * Return the thumbnail file path
     * @return string
     */
    public function getThumbnailPath()
    {
        return $this->thumbnail;
    }
    
    /**
     * Define the path for the thumbnail file / preview picture
     * @param string $value
     * @return \PhpOffice\PhpPresentation\PresentationProperties
     */
    public function setThumbnailPath($path = '')
    {
        if (file_exists($path)) {
            $this->thumbnail = $path;
        }
        return $this;
    }

    /**
     * Mark a document as final
     * @param bool $state
     * @return PhpPresentation
     */
    public function markAsFinal($state = true)
    {
        if (is_bool($state)) {
            $this->markAsFinal = $state;
        }
        return $this;
    }

    /**
     * Return if this document is marked as final
     * @return bool
     */
    public function isMarkedAsFinal()
    {
        return $this->markAsFinal;
    }

    /**
     * Set the zoom of the document (in percentage)
     * @param float $zoom
     * @return PhpPresentation
     */
    public function setZoom($zoom = 1)
    {
        if (is_numeric($zoom)) {
            $this->zoom = $zoom;
        }
        return $this;
    }

    /**
     * Return the zoom (in percentage)
     * @return float
     */
    public function getZoom()
    {
        return $this->zoom;
    }
}
