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

namespace PhpOffice\PhpPresentation;

use PhpOffice\PhpPresentation\Exception\ShapeContainerAlreadyAssignedException;
use PhpOffice\PhpPresentation\Shape\Group;
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
     * @var null|ShapeContainerInterface
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
     * @var null|Fill
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
    protected $rotation = 0;

    /**
     * Shadow.
     *
     * @var null|Shadow
     */
    protected $shadow;

    /**
     * @var null|Hyperlink
     */
    protected $hyperlink;

    /**
     * @var null|Placeholder
     */
    protected $placeholder;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Name.
     *
     * @var string
     */
    protected $name = '';

    /**
     * Create a new self.
     */
    public function __construct()
    {
        $this->offsetX = $this->offsetY = $this->width = $this->height = 0;
        $this->fill = new Fill();
        $this->shadow = new Shadow();
        $this->border = new Border();

        $this->border->setLineStyle(Border::LINE_NONE);
    }

    /**
     * Magic Method : clone.
     */
    public function __clone()
    {
        $this->container = null;
        $this->name = $this->name;
        $this->border = clone $this->border;
        if (isset($this->fill)) {
            $this->fill = clone $this->fill;
        }
        if (isset($this->shadow)) {
            $this->shadow = clone $this->shadow;
        }
        if (isset($this->placeholder)) {
            $this->placeholder = clone $this->placeholder;
        }
        if (isset($this->hyperlink)) {
            $this->hyperlink = clone $this->hyperlink;
        }
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
     * @param bool $pOverrideOld If a Slide has already been assigned, overwrite it and remove image from old Slide?
     *
     * @return $this
     */
    public function setContainer(?ShapeContainerInterface $pValue = null, $pOverrideOld = false)
    {
        if (null === $this->container) {
            // Add drawing to ShapeContainerInterface
            $this->container = $pValue;
            if (null !== $this->container) {
                $this->container->addShape($this);
            }
        } else {
            if ($pOverrideOld) {
                // Remove drawing from old ShapeContainerInterface
                foreach ($this->container->getShapeCollection() as $key => $shape) {
                    if ($shape->getHashCode() == $this->getHashCode()) {
                        $this->container->unsetShape($key);
                        $this->container = null;

                        break;
                    }
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
     * Get Name.
     */
    public function getName(): string
    {
        return $this->name;
    }

    /**
     * Set Name.
     *
     * @return static
     */
    public function setName(string $pValue = ''): self
    {
        $this->name = $pValue;

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
     * Set OffsetX (in pixels).
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
     * Get Width (in pixels).
     *
     * @return int
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set Width (in pixels).
     *
     * @return $this
     */
    public function setWidth(int $pValue = 0)
    {
        $this->width = $pValue;

        return $this;
    }

    /**
     * Get Height (in pixels).
     *
     * @return int
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Set Height (in pixels).
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
     */
    public function getRotation(): int
    {
        return $this->rotation;
    }

    /**
     * Set Rotation.
     */
    public function setRotation(int $pValue = 0): self
    {
        $this->rotation = $pValue;

        return $this;
    }

    public function getFill(): ?Fill
    {
        return $this->fill;
    }

    public function setFill(?Fill $pValue = null): self
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
    public function setShadow(?Shadow $pValue = null)
    {
        $this->shadow = $pValue;

        return $this;
    }

    /**
     * Has Hyperlink?
     */
    public function hasHyperlink(): bool
    {
        return null !== $this->hyperlink;
    }

    /**
     * Get Hyperlink.
     */
    public function getHyperlink(): Hyperlink
    {
        if (null === $this->hyperlink) {
            $this->hyperlink = new Hyperlink();
        }

        return $this->hyperlink;
    }

    /**
     * Set Hyperlink.
     */
    public function setHyperlink(?Hyperlink $pHyperlink = null): self
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
        return md5((is_object($this->container) ? $this->container->getHashCode() : '') . $this->offsetX . $this->offsetY . $this->width . $this->height . $this->rotation . (null === $this->getFill() ? '' : $this->getFill()->getHashCode()) . (null === $this->shadow ? '' : $this->shadow->getHashCode()) . (null === $this->hyperlink ? '' : $this->hyperlink->getHashCode()) . __CLASS__);
    }

    /**
     * Get hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return null|int Hash index
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
        return null !== $this->placeholder;
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
