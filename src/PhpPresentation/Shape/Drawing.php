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

namespace PhpOffice\PhpPresentation\Shape;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * Drawing element
 */
class Drawing extends AbstractDrawing implements ComparableInterface
{
    /**
     * Path
     *
     * @var string
     */
    private $path;

    /**
     * Create a new \PhpOffice\PhpPresentation\Slide\Drawing
     */
    public function __construct()
    {
        // Initialise values
        $this->path = '';

        // Initialize parent
        parent::__construct();
    }

    /**
     * Get Filename
     *
     * @return string
     */
    public function getFilename()
    {
        return basename($this->path);
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename()
    {
        return str_replace('.' . $this->getExtension(), '', $this->getFilename()) . $this->getImageIndex() . '.' . $this->getExtension();
    }

    /**
     * Get Extension
     *
     * @return string
     */
    public function getExtension()
    {
        $exploded = explode(".", basename($this->path));

        return $exploded[count($exploded) - 1];
    }

    /**
     * Get Path
     *
     * @return string
     */
    public function getPath()
    {
        return $this->path;
    }

    /**
     * Set Path
     *
     * @param  string                      $pValue      File path
     * @param  boolean                     $pVerifyFile Verify file
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\Drawing
     */
    public function setPath($pValue = '', $pVerifyFile = true)
    {
        if ($pVerifyFile) {
            if (file_exists($pValue)) {
                $this->path = $pValue;

                if ($this->width == 0 && $this->height == 0) {
                    // Get width/height
                    list($this->width, $this->height) = getimagesize($pValue);
                }
            } else {
                throw new \Exception("File $pValue not found!");
            }
        } else {
            $this->path = $pValue;
        }

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->path . parent::getHashCode() . __CLASS__);
    }
}
