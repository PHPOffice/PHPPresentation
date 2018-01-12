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

namespace PhpOffice\PhpPresentation\Slide;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\GeometryCalculator;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing\File;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Traits\ShapeCollection;

abstract class AbstractSlide implements ComparableInterface, ShapeContainerInterface
{
    use ShapeCollection;

    /**
     * @var string
     */
    protected $relsIndex;

    /**
     * @var null|Transition
     */
    protected $slideTransition;

    /**
     * Extent Y.
     *
     * @var int
     */
    protected $extentY;

    /**
     * Extent X.
     *
     * @var int
     */
    protected $extentX;

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
     * Slide identifier.
     *
     * @var string
     */
    protected $identifier;

    /**
     * Hash index.
     *
     * @var int
     */
    protected $hashIndex;

    /**
     * Parent presentation.
     *
     * @var null|PhpPresentation
     */
    protected $parent;

    /**
     * Background of the slide.
     *
     * @var AbstractBackground
     */
    protected $background;

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
     * Create line shape.
     *
     * @param int $fromX Starting point x offset
     * @param int $fromY Starting point y offset
     * @param int $toX Ending point x offset
     * @param int $toY Ending point y offset
     */
    public function createLineShape(int $fromX, int $fromY, int $toX, int $toY): Line
    {
        $shape = new Line($fromX, $fromY, $toX, $toY);
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create geometric shape.
     */
    public function createAutoShape(): AutoShape
    {
        $shape = new AutoShape();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create chart shape.
     */
    public function createChartShape(): Chart
    {
        $shape = new Chart();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create drawing shape.
     */
    public function createDrawingShape(): File
    {
        $shape = new File();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create table shape.
     *
     * @param int $columns Number of columns
     */
    public function createTableShape(int $columns = 1): Table
    {
        $shape = new Table($columns);
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Creates a group within this slide.
     */
    public function createGroup(): Group
    {
        $shape = new Group();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Get parent.
     */
    public function getParent(): ?PhpPresentation
    {
        return $this->parent;
    }

    /**
     * Re-bind parent.
     */
    public function rebindParent(PhpPresentation $parent): self
    {
        $this->parent->removeSlideByIndex($this->parent->getIndex($this));
        $this->parent = $parent;

        return $this;
    }

    public function getBackground(): ?AbstractBackground
    {
        return $this->background;
    }

    public function setBackground(?AbstractBackground $background = null): self
    {
        $this->background = $background;

        return $this;
    }

    public function getTransition(): ?Transition
    {
        return $this->slideTransition;
    }

    public function setTransition(?Transition $transition = null): self
    {
        $this->slideTransition = $transition;

        return $this;
    }

    public function getRelsIndex(): string
    {
        return $this->relsIndex;
    }

    public function setRelsIndex(string $indexName): self
    {
        $this->relsIndex = $indexName;

        return $this;
    }
}
