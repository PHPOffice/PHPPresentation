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

namespace PhpOffice\PhpPresentation\Shape;

use ArrayObject;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\GeometryCalculator;
use PhpOffice\PhpPresentation\ShapeContainerInterface;

class Group extends AbstractShape implements ShapeContainerInterface
{
    /**
     * Collection of shapes.
     *
     * @var array<int, AbstractShape>|ArrayObject<int, AbstractShape>
     */
    private $shapeCollection;

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

    public function __construct()
    {
        parent::__construct();

        // Shape collection
        $this->shapeCollection = new ArrayObject();
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
     * Get X Offset.
     */
    public function getOffsetX(): int
    {
        if (empty($this->offsetX)) {
            $offsets = GeometryCalculator::calculateOffsets($this);
            $this->offsetX = $offsets[GeometryCalculator::X];
            $this->offsetY = $offsets[GeometryCalculator::Y];
        }

        return $this->offsetX;
    }

    /**
     * Ignores setting the X Offset, preserving the default behavior.
     *
     * @return $this
     */
    public function setOffsetX(int $pValue = 0)
    {
        return $this;
    }

    /**
     * Get Y Offset.
     */
    public function getOffsetY(): int
    {
        if (empty($this->offsetY)) {
            $offsets = GeometryCalculator::calculateOffsets($this);
            $this->offsetX = $offsets[GeometryCalculator::X];
            $this->offsetY = $offsets[GeometryCalculator::Y];
        }

        return $this->offsetY;
    }

    /**
     * Ignores setting the Y Offset, preserving the default behavior.
     *
     * @return $this
     */
    public function setOffsetY(int $pValue = 0)
    {
        return $this;
    }

    /**
     * Get X Extent.
     */
    public function getExtentX(): int
    {
        if (null === $this->extentX) {
            $extents = GeometryCalculator::calculateExtents($this);
            $this->extentX = $extents[GeometryCalculator::X] - $this->getOffsetX();
            $this->extentY = $extents[GeometryCalculator::Y] - $this->getOffsetY();
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
            $this->extentX = $extents[GeometryCalculator::X] - $this->getOffsetX();
            $this->extentY = $extents[GeometryCalculator::Y] - $this->getOffsetY();
        }

        return $this->extentY;
    }

    /**
     * Ignores setting the width, preserving the default behavior.
     *
     * @return self
     */
    public function setWidth(int $pValue = 0)
    {
        return $this;
    }

    /**
     * Ignores setting the height, preserving the default behavior.
     *
     * @return $this
     */
    public function setHeight(int $pValue = 0)
    {
        return $this;
    }

    /**
     * Create rich text shape.
     *
     * @return RichText
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
     *
     * @return Line
     */
    public function createLineShape(int $fromX, int $fromY, int $toX, int $toY): Line
    {
        $shape = new Line($fromX, $fromY, $toX, $toY);
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create chart shape.
     *
     * @return Chart
     */
    public function createChartShape(): Chart
    {
        $shape = new Chart();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create drawing shape.
     *
     * @return Drawing\File
     */
    public function createDrawingShape(): Drawing\File
    {
        $shape = new Drawing\File();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create table shape.
     *
     * @param int $columns Number of columns
     *
     * @return Table
     */
    public function createTableShape(int $columns = 1): Table
    {
        $shape = new Table($columns);
        $this->addShape($shape);

        return $shape;
    }
}
