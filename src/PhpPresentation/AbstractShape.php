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

use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;

/**
 * Abstract shape
 */
abstract class AbstractShape implements ComparableInterface
{
    /**
     * Container
     *
     * @var \PhpOffice\PhpPresentation\ShapeContainerInterface
     */
    protected $container;

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
     * @var \PhpOffice\PhpPresentation\Style\Fill
     */
    private $fill;

    /**
     * Border
     *
     * @var \PhpOffice\PhpPresentation\Style\Border
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
     * @var \PhpOffice\PhpPresentation\Style\Shadow
     */
    protected $shadow;

    /**
     * Hyperlink
     *
     * @var \PhpOffice\PhpPresentation\Shape\Hyperlink
     */
    protected $hyperlink;

    /**
     * PlaceHolder
     * @var \PhpOffice\PhpPresentation\Shape\Placeholder
     */
    protected $placeholder;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new self
     */
    public function __construct()
    {
        // Initialise values
        $this->container = null;
        $this->offsetX = 0;
        $this->offsetY = 0;
        $this->width = 0;
        $this->height = 0;
        $this->rotation = 0;
        $this->fill = new Style\Fill();
        $this->border = new Style\Border();
        $this->shadow = new Style\Shadow();

        $this->border->setLineStyle(Style\Border::LINE_NONE);
    }

    /**
     * Magic Method : clone
     */
    public function __clone()
    {
        $this->container = null;
        $this->fill = clone $this->fill;
        $this->border = clone $this->border;
        $this->shadow = clone $this->shadow;
    }

    /**
     * Get Container, Slide or Group
     *
     * @return \PhpOffice\PhpPresentation\Container
     */
    public function getContainer()
    {
        return $this->container;
    }

    /**
     * Set Container, Slide or Group
     *
     * @param  \PhpOffice\PhpPresentation\ShapeContainerInterface $pValue
     * @param  bool $pOverrideOld If a Slide has already been assigned, overwrite it and remove image from old Slide?
     * @throws \Exception
     * @return self
     */
    public function setContainer(ShapeContainerInterface $pValue = null, $pOverrideOld = false)
    {
        if (is_null($this->container)) {
            // Add drawing to \PhpOffice\PhpPresentation\ShapeContainerInterface
            $this->container = $pValue;
            if (!is_null($this->container)) {
                $this->container->getShapeCollection()->append($this);
            }
        } else {
            if ($pOverrideOld) {
                // Remove drawing from old \PhpOffice\PhpPresentation\ShapeContainerInterface
                $iterator = $this->container->getShapeCollection()->getIterator();

                while ($iterator->valid()) {
                    if ($iterator->current()->getHashCode() == $this->getHashCode()) {
                        $this->container->getShapeCollection()->offsetUnset($iterator->key());
                        $this->container = null;
                        break;
                    }
                    $iterator->next();
                }

                // Set new \PhpOffice\PhpPresentation\Slide
                $this->setContainer($pValue);
            } else {
                throw new \Exception("A \PhpOffice\PhpPresentation\ShapeContainerInterface has already been assigned. Shapes can only exist on one \PhpOffice\PhpPresentation\ShapeContainerInterface.");
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
     * @param  int $pValue
     * @return self
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
     * @param  int $pValue
     * @return self
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
     * @param  int $pValue
     * @return self
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
     * @param  int $pValue
     * @return self
     */
    public function setHeight($pValue = 0)
    {
        $this->height = $pValue;
        return $this;
    }

    /**
     * Set width and height with proportional resize
     *
     * @param  int $width
     * @param  int $height
     * @example $objDrawing->setWidthAndHeight(160,120);
     * @return self
     */
    public function setWidthAndHeight($width = 0, $height = 0)
    {
        $this->width = $width;
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
     * @param  int $pValue
     * @return self
     */
    public function setRotation($pValue = 0)
    {
        $this->rotation = $pValue;
        return $this;
    }

    /**
     * Get Fill
     *
     * @return \PhpOffice\PhpPresentation\Style\Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * Set Fill
     * @param \PhpOffice\PhpPresentation\Style\Fill $pValue
     * @return \PhpOffice\PhpPresentation\AbstractShape
     */
    public function setFill(Fill $pValue = null)
    {
        $this->fill = $pValue;
        return $this;
    }

    /**
     * Get Border
     *
     * @return \PhpOffice\PhpPresentation\Style\Border
     */
    public function getBorder()
    {
        return $this->border;
    }

    /**
     * Get Shadow
     *
     * @return \PhpOffice\PhpPresentation\Style\Shadow
     */
    public function getShadow()
    {
        return $this->shadow;
    }

    /**
     * Set Shadow
     *
     * @param  \PhpOffice\PhpPresentation\Style\Shadow $pValue
     * @throws \Exception
     * @return self
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
     * @return \PhpOffice\PhpPresentation\Shape\Hyperlink
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
     * @param  \PhpOffice\PhpPresentation\Shape\Hyperlink $pHyperlink
     * @throws \Exception
     * @return self
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
        return md5((is_object($this->container) ? $this->container->getHashCode() : '') . $this->offsetX . $this->offsetY . $this->width . $this->height . $this->rotation . (is_null($this->getFill()) ? '' : $this->getFill()->getHashCode()) . (is_null($this->shadow) ? '' : $this->shadow->getHashCode()) . (is_null($this->hyperlink) ? '' : $this->hyperlink->getHashCode()) . __CLASS__);
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

    public function isPlaceholder()
    {
        return !is_null($this->placeholder);
    }

    public function getPlaceholder()
    {
        if (!$this->isPlaceholder()) {
            return null;
        }
        return $this->placeholder;
    }

    /**
     * @param \PhpOffice\PhpPresentation\Shape\Placeholder $placeholder
     * @return $this
     */
    public function setPlaceHolder(Placeholder $placeholder)
    {
        $this->placeholder = $placeholder;
        return $this;
    }
}
