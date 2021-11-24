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

namespace PhpOffice\PhpPresentation\Slide;

use ArrayObject;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\GeometryCalculator;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Slide;

class Note implements ComparableInterface, ShapeContainerInterface
{
    /**
     * Parent slide.
     *
     * @var Slide
     */
    private $parent;

    /**
     * Collection of shapes.
     *
     * @var array<int, AbstractShape>|ArrayObject<int, AbstractShape>
     */
    private $shapeCollection;

    /**
     * Note identifier.
     *
     * @var string
     */
    private $identifier;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

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
     * Extent X.
     *
     * @var int
     */
    protected $extentX;

    /**
     * Extent Y.
     *
     * @var int
     */
    protected $extentY;

    /**
     * Create a new note.
     *
     * @param Slide $pParent
     */
    public function __construct(Slide $pParent = null)
    {
        // Set parent
        $this->parent = $pParent;

        // Shape collection
        $this->shapeCollection = new ArrayObject();

        // Set identifier
        $this->identifier = md5(rand(0, 9999) . time());
    }

    /**
     * Get collection of shapes.
     *
     * @return array<int, AbstractShape>|ArrayObject<int, AbstractShape>
     */
    public function getShapeCollection()
    {
        return $this->shapeCollection;
    }

    /**
     * Add shape to slide.
     *
     * @return AbstractShape
     */
    public function addShape(AbstractShape $shape): AbstractShape
    {
        $shape->setContainer($this);

        return $shape;
    }

    /**
     * Create rich text shape.
     */
    public function createRichTextShape(): RichText
    {
        $shape = new RichText();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Get parent.
     *
     * @return Slide
     */
    public function getParent()
    {
        return $this->parent;
    }

    /**
     * Set parent.
     *
     * @return Note
     */
    public function setParent(Slide $parent)
    {
        $this->parent = $parent;

        return $this;
    }

    /**
     * Get X Offset.
     */
    public function getOffsetX(): int
    {
        if (null === $this->offsetX) {
            $offsets = GeometryCalculator::calculateOffsets($this);
            $this->offsetX = $offsets[GeometryCalculator::X];
            $this->offsetY = $offsets[GeometryCalculator::Y];
        }

        return $this->offsetX;
    }

    /**
     * Get Y Offset.
     */
    public function getOffsetY(): int
    {
        if (null === $this->offsetY) {
            $offsets = GeometryCalculator::calculateOffsets($this);
            $this->offsetX = $offsets[GeometryCalculator::X];
            $this->offsetY = $offsets[GeometryCalculator::Y];
        }

        return $this->offsetY;
    }

    /**
     * Get X Extent.
     */
    public function getExtentX(): int
    {
        if (null === $this->extentX) {
            $extents = GeometryCalculator::calculateExtents($this);
            $this->extentX = $extents[GeometryCalculator::X];
            $this->extentY = $extents[GeometryCalculator::Y];
        }

        return $this->extentX;
    }

    /**
     * Get Y Extent.
     */
    public function getExtentY(): int
    {
        if (null === $this->extentY) {
            $extents = GeometryCalculator::calculateExtents($this);
            $this->extentX = $extents[GeometryCalculator::X];
            $this->extentY = $extents[GeometryCalculator::Y];
        }

        return $this->extentY;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5($this->identifier . __CLASS__);
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
}
