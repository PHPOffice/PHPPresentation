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

namespace PhpOffice\PhpPowerpoint;

use PhpOffice\PhpPowerpoint\IComparable;
use PhpOffice\PhpPowerpoint\Shape\Hyperlink;
use PhpOffice\PhpPowerpoint\Style\Shadow;

/**
 * PHPPowerPoint_Shape
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
abstract class Shape implements IComparable
{
    /**
     * Slide
     *
     * @var PHPPowerPoint_Slide
     */
    protected $slide;

    /**
     * Offset X
     *
     * @var int
     */
    protected $offsetX;

    /**
     * Offset Y
     *
     * @var int
     */
    protected $offsetY;

    /**
     * Width
     *
     * @var int
     */
    protected $width;

    /**
     * Height
     *
     * @var int
     */
    protected $height;

    /**
     * Fill
     *
     * @var PHPPowerPoint_Style_Fill
     */
    private $fill;

    /**
     * Border
     *
     * @var PHPPowerPoint_Style_Border
     */
    private $border;

    /**
     * Rotation
     *
     * @var int
     */
    protected $rotation;

    /**
     * Shadow
     *
     * @var PHPPowerPoint_Style_Shadow
     */
    protected $shadow;

    /**
     * Hyperlink
     *
     * @var PHPPowerPoint_Shape_Hyperlink
     */
    protected $hyperlink;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;
    
    /**
     * Create a new PHPPowerPoint_Shape
     */
    public function __construct()
    {
        // Initialise values
        $this->slide    = null;
        $this->offsetX  = 0;
        $this->offsetY  = 0;
        $this->width    = 0;
        $this->height   = 0;
        $this->rotation = 0;
        $this->fill     = new Style\Fill();
        $this->border   = new Style\Border();
        $this->shadow   = new Style\Shadow();

        $this->border->setLineStyle(Style\Border::LINE_NONE);
    }

    /**
     * Get Slide
     *
     * @return PHPPowerPoint_Slide
     */
    public function getSlide()
    {
        return $this->slide;
    }

    /**
     * Set Slide
     *
     * @param  PHPPowerPoint_Slide $pValue
     * @param  bool                $pOverrideOld If a Slide has already been assigned, overwrite it and remove image from old Slide?
     * @throws \Exception
     * @return PHPPowerPoint_Shape
     */
    public function setSlide(Slide $pValue = null, $pOverrideOld = false)
    {
        if (is_null($this->slide)) {
            // Add drawing to PHPPowerPoint_Slide
            $this->slide = $pValue;
            $this->slide->getShapeCollection()->append($this);
        } else {
            if ($pOverrideOld) {
                // Remove drawing from old PHPPowerPoint_Slide
                $iterator = $this->slide->getShapeCollection()->getIterator();

                while ($iterator->valid()) {
                    if ($iterator->current()->getHashCode() == $this->getHashCode()) {
                        $this->slide->getShapeCollection()->offsetUnset($iterator->key());
                        $this->slide = null;
                        break;
                    }
                }

                // Set new PHPPowerPoint_Slide
                $this->setSlide($pValue);
            } else {
                throw new \Exception("A PHPPowerPoint_Slide has already been assigned. Shapes can only exist on one PHPPowerPoint_Slide.");
            }
        }

        return $this;
    }

    /**
     * Get OffsetX
     *
     * @return int
     */
    public function getOffsetX()
    {
        return $this->offsetX;
    }

    /**
     * Set OffsetX
     *
     * @param  int                 $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setOffsetX($pValue = 0)
    {
        $this->offsetX = $pValue;

        return $this;
    }

    /**
     * Get OffsetY
     *
     * @return int
     */
    public function getOffsetY()
    {
        return $this->offsetY;
    }

    /**
     * Set OffsetY
     *
     * @param  int                 $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setOffsetY($pValue = 0)
    {
        $this->offsetY = $pValue;

        return $this;
    }

    /**
     * Get Width
     *
     * @return int
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set Width
     *
     * @param  int                 $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setWidth($pValue = 0)
    {
        $this->width = $pValue;
        return $this;
    }

    /**
     * Get Height
     *
     * @return int
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Set Height
     *
     * @param  int                 $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setHeight($pValue = 0)
    {
        $this->height = $pValue;
        return $this;
    }

    /**
     * Set width and height with proportional resize
     *
     * @param  int                 $width
     * @param  int                 $height
     * @example $objDrawing->setWidthAndHeight(160,120);
     * @return PHPPowerPoint_Shape
     */
    public function setWidthAndHeight($width = 0, $height = 0)
    {
        $this->width  = $width;
        $this->height = $height;
        return $this;
    }

    /**
     * Get Rotation
     *
     * @return int
     */
    public function getRotation()
    {
        return $this->rotation;
    }

    /**
     * Set Rotation
     *
     * @param  int                 $pValue
     * @return PHPPowerPoint_Shape
     */
    public function setRotation($pValue = 0)
    {
        $this->rotation = $pValue;
        return $this;
    }

    /**
     * Get Fill
     *
     * @return PHPPowerPoint_Style_Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Get Border
     *
     * @return PHPPowerPoint_Style_Border
     */
    public function getBorder()
    {
        return $this->border;
    }

    /**
     * Get Shadow
     *
     * @return PHPPowerPoint_Style_Shadow
     */
    public function getShadow()
    {
        return $this->shadow;
    }

    /**
     * Set Shadow
     *
     * @param  PHPPowerPoint_Style_Shadow $pValue
     * @throws \Exception
     * @return PHPPowerPoint_Shape
     */
    public function setShadow(Shadow $pValue = null)
    {
        $this->shadow = $pValue;
        return $this;
    }

    /**
     * Has Hyperlink?
     *
     * @return boolean
     */
    public function hasHyperlink()
    {
        return !is_null($this->hyperlink);
    }

    /**
     * Get Hyperlink
     *
     * @return PHPPowerPoint_Shape_Hyperlink
     */
    public function getHyperlink()
    {
        if (is_null($this->hyperlink)) {
            $this->hyperlink = new Hyperlink();
        }
        return $this->hyperlink;
    }

    /**
     * Set Hyperlink
     *
     * @param  PHPPowerPoint_Shape_Hyperlink $pHyperlink
     * @throws \Exception
     * @return PHPPowerPoint_Shape
     */
    public function setHyperlink(Hyperlink $pHyperlink = null)
    {
        $this->hyperlink = $pHyperlink;
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->slide->getHashCode() . $this->offsetX . $this->offsetY . $this->width . $this->height . $this->rotation . $this->getFill()->getHashCode() . $this->shadow->getHashCode() . (is_null($this->hyperlink) ? '' : $this->hyperlink->getHashCode()) . __CLASS__);
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
