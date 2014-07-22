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

namespace PhpOffice\PhpPowerpoint\Shape\Chart;

use PhpOffice\PhpPowerpoint\ComparableInterface;

/**
 * \PhpOffice\PhpPowerpoint\Shape\Chart\View3D
 */
class View3D implements ComparableInterface
{
    /**
     * Rotation X
     *
     * @var int
     */
    protected $rotationX = 0;

    /**
     * Rotation Y
     *
     * @var int
     */
    protected $rotationY = 0;

    /**
     * Right Angle Axes
     *
     * @var boolean
     */
    private $rightAngleAxes = true;

    /**
     * Perspective
     *
     * @var int
     */
    private $perspective = 30;

    /**
     * Height Percent
     *
     * @var int
     */
    private $heightPercent = 100;

    /**
     * Depth Percent
     *
     * @var int
     */
    private $depthPercent = 100;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Shape\Chart\View3D instance
     */
    public function __construct()
    {
    }

    /**
     * Get Rotation X
     *
     * @return int
     */
    public function getRotationX()
    {
        return $this->rotationX;
    }

    /**
     * Set Rotation X (-90 to 90)
     *
     * @param  int                              $pValue
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\View3D
     */
    public function setRotationX($pValue = 0)
    {
        $this->rotationX = $pValue;

        return $this;
    }

    /**
     * Get Rotation Y
     *
     * @return int
     */
    public function getRotationY()
    {
        return $this->rotationY;
    }

    /**
     * Set Rotation Y (-90 to 90)
     *
     * @param  int                              $pValue
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\View3D
     */
    public function setRotationY($pValue = 0)
    {
        $this->rotationY = $pValue;

        return $this;
    }

    /**
     * Get RightAngleAxes
     *
     * @return boolean
     */
    public function hasRightAngleAxes()
    {
        return $this->rightAngleAxes;
    }

    /**
     * Set RightAngleAxes
     *
     * @param  boolean                          $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\View3D
     */
    public function setRightAngleAxes($value = true)
    {
        $this->rightAngleAxes = $value;

        return $this;
    }

    /**
     * Get Perspective
     *
     * @return int
     */
    public function getPerspective()
    {
        return $this->perspective;
    }

    /**
     * Set Perspective (0 to 100)
     *
     * @param  int                              $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\View3D
     */
    public function setPerspective($value = 30)
    {
        $this->perspective = $value;

        return $this;
    }

    /**
     * Get HeightPercent
     *
     * @return int
     */
    public function getHeightPercent()
    {
        return $this->heightPercent;
    }

    /**
     * Set HeightPercent (5 to 500)
     *
     * @param  int  $value
     * @return self
     */
    public function setHeightPercent($value = 100)
    {
        $this->heightPercent = $value;

        return $this;
    }

    /**
     * Get DepthPercent
     *
     * @return int
     */
    public function getDepthPercent()
    {
        return $this->depthPercent;
    }

    /**
     * Set DepthPercent (20 to 2000)
     *
     * @param  int  $value
     * @return self
     */
    public function setDepthPercent($value = 100)
    {
        $this->depthPercent = $value;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->rotationX . $this->rotationY . ($this->rightAngleAxes ? 't' : 'f') . $this->perspective . $this->heightPercent . $this->depthPercent . __CLASS__);
    }

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
        return $this->hashIndex;
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
        $this->hashIndex = $value;
        return $this;
    }
}
