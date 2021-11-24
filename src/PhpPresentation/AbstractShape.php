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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation;

use PhpOffice\PhpPresentation\Exception\ShapeContainerAlreadyAssignedException;
use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;

/**
 * Abstract shape.
 */
abstract class AbstractShape implements ComparableInterface
{
    /**
     * Container.
     *
     * @var ShapeContainerInterface|null
     */
    protected $container;

    /**
     * Offset X.
     *
     * @var int
     */
    protected $offsetX;

    /**
     * Offset Y.
     *
     * @var int
     */
    protected $offsetY;

    /**
     * Width.
     *
     * @var int
     */
    protected $width;

    /**
     * Height.
     *
     * @var int
     */
    protected $height;

    /**
     * @var Fill|null
     */
    private $fill;

    /**
     * Border.
     *
     * @var Border
     */
    private $border;

    /**
     * Rotation.
     *
     * @var int
     */
    protected $rotation;

    /**
     * Shadow.
     *
     * @var Shadow|null
     */
    protected $shadow;

    /**
     * @var Hyperlink|null
     */
    protected $hyperlink;

    /**
     * @var Placeholder|null
     */
    protected $placeholder;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new self.
     */
    public function __construct()
    {
        $this->offsetX = $this->offsetY = $this->width = $this->height = $this->rotation = 0;
        $this->fill = new Fill();
        $this->shadow = new Shadow();
        $this->border = new Border();
        $this->border->setLineStyle(Style\Border::LINE_NONE);
    }

    /**
     * Magic Method : clone.
     */
    public function __clone()
    {
        $this->container = null;
        $this->fill = clone $this->fill;
        $this->border = clone $this->border;
        $this->shadow = clone $this->shadow;
    }

    /**
     * Get Container, Slide or Group.
     */
    public function getContainer(): ?ShapeContainerInterface
    {
        return $this->container;
    }

    /**
     * Set Container, Slide or Group.
     *
     * @param ShapeContainerInterface $pValue
     * @param bool $pOverrideOld If a Slide has already been assigned, overwrite it and remove image from old Slide?
     *
     * @throws ShapeContainerAlreadyAssignedException
     *
     * @return $this
     */
    public function setContainer(ShapeContainerInterface $pValue = null, $pOverrideOld = false)
    {
        if (is_null($this->container)) {
            // Add drawing to ShapeContainerInterface
            $this->container = $pValue;
            if (!is_null($this->container)) {
                $this->container->getShapeCollection()->append($this);
            }
        } else {
            if ($pOverrideOld) {
                // Remove drawing from old ShapeContainerInterface
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
                throw new ShapeContainerAlreadyAssignedException(self::class);
            }
        }

        return $this;
    }

    /**
     * Get OffsetX.
     */
    public function getOffsetX(): int
    {
        return $this->offsetX;
    }

    /**
     * Set OffsetX.
     *
     * @return $this
     */
    public function setOffsetX(int $pValue = 0)
    {
        $this->offsetX = $pValue;

        return $this;
    }

    /**
     * Get OffsetY.
     *
     * @return int
     */
    public function getOffsetY()
    {
        return $this->offsetY;
    }

    /**
     * Set OffsetY.
     *
     * @return $this
     */
    public function setOffsetY(int $pValue = 0)
    {
        $this->offsetY = $pValue;

        return $this;
    }

    /**
     * Get Width.
     *
     * @return int
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set Width.
     *
     * @return $this
     */
    public function setWidth(int $pValue = 0)
    {
        $this->width = $pValue;

        return $this;
    }

    /**
     * Get Height.
     *
     * @return int
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Set Height.
     *
     * @return $this
     */
    public function setHeight(int $pValue = 0)
    {
        $this->height = $pValue;

        return $this;
    }

    /**
     * Set width and height with proportional resize.
     *
     * @return self
     */
    public function setWidthAndHeight(int $width = 0, int $height = 0)
    {
        $this->width = $width;
        $this->height = $height;

        return $this;
    }

    /**
     * Get Rotation.
     *
     * @return int
     */
    public function getRotation()
    {
        return $this->rotation;
    }

    /**
     * Set Rotation.
     *
     * @param int $pValue
     *
     * @return $this
     */
    public function setRotation($pValue = 0)
    {
        $this->rotation = $pValue;

        return $this;
    }

    public function getFill(): ?Fill
    {
        return $this->fill;
    }

    public function setFill(Fill $pValue = null): self
    {
        $this->fill = $pValue;

        return $this;
    }

    public function getBorder(): Border
    {
        return $this->border;
    }

    public function getShadow(): ?Shadow
    {
        return $this->shadow;
    }

    /**
     * @return $this
     */
    public function setShadow(Shadow $pValue = null)
    {
        $this->shadow = $pValue;

        return $this;
    }

    /**
     * Has Hyperlink?
     *
     * @return bool
     */
    public function hasHyperlink()
    {
        return !is_null($this->hyperlink);
    }

    /**
     * Get Hyperlink
     */
    public function getHyperlink(): Hyperlink
    {
        if (is_null($this->hyperlink)) {
            $this->hyperlink = new Hyperlink();
        }

        return $this->hyperlink;
    }

    /**
     * Set Hyperlink
     */
    public function setHyperlink(Hyperlink $pHyperlink = null): self
    {
        $this->hyperlink = $pHyperlink;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5((is_object($this->container) ? $this->container->getHashCode() : '') . $this->offsetX . $this->offsetY . $this->width . $this->height . $this->rotation . (is_null($this->getFill()) ? '' : $this->getFill()->getHashCode()) . (is_null($this->shadow) ? '' : $this->shadow->getHashCode()) . (is_null($this->hyperlink) ? '' : $this->hyperlink->getHashCode()) . __CLASS__);
    }

    /**
     * Get hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return int|null Hash index
     */
    public function getHashIndex(): ?int
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param int $value Hash index
     *
     * @return $this
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;

        return $this;
    }

    public function isPlaceholder(): bool
    {
        return !is_null($this->placeholder);
    }

    public function getPlaceholder(): ?Placeholder
    {
        if (!$this->isPlaceholder()) {
            return null;
        }

        return $this->placeholder;
    }

    public function setPlaceHolder(Placeholder $placeholder): self
    {
        $this->placeholder = $placeholder;

        return $this;
    }
}
