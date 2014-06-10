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

use PhpOffice\PhpPowerpoint\IComparable;

/**
 * PHPPowerPoint_Shape_Drawing
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Drawing extends BaseDrawing implements IComparable
{
    /**
     * Path
     *
     * @var string
     */
    private $_path;

    /**
     * Create a new PHPPowerPoint_Slide_Drawing
     */
    public function __construct()
    {
        // Initialise values
        $this->_path = '';

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
        return basename($this->_path);
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
        $exploded = explode(".", basename($this->_path));

        return $exploded[count($exploded) - 1];
    }

    /**
     * Get Path
     *
     * @return string
     */
    public function getPath()
    {
        return $this->_path;
    }

    /**
     * Set Path
     *
     * @param  string                      $pValue      File path
     * @param  boolean                     $pVerifyFile Verify file
     * @throws Exception
     * @return PHPPowerPoint_Shape_Drawing
     */
    public function setPath($pValue = '', $pVerifyFile = true)
    {
        if ($pVerifyFile) {
            if (file_exists($pValue)) {
                $this->_path = $pValue;

                if ($this->_width == 0 && $this->_height == 0) {
                    // Get width/height
                    list($this->_width, $this->_height) = getimagesize($pValue);
                }
            } else {
                throw new Exception("File $pValue not found!");
            }
        } else {
            $this->_path = $pValue;
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
        return md5($this->_path . parent::getHashCode() . __CLASS__);
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
