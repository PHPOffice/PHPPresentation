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
namespace PhpOffice\PhpPresentation\Slide;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\GeometryCalculator;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing\File;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Slide;

abstract class AbstractSlide implements ComparableInterface, ShapeContainerInterface
{
    /**
     * @var string
     */
    protected $relsIndex;
    /**
     *
     * @var \PhpOffice\PhpPresentation\Slide\Transition
     */
    protected $slideTransition;

    /**
     * Collection of shapes
     *
     * @var \ArrayObject|\PhpOffice\PhpPresentation\AbstractShape[]
     */
    protected $shapeCollection = null;
    /**
     * Extent Y
     *
     * @var int
     */
    protected $extentY;
    /**
     * Extent X
     *
     * @var int
     */
    protected $extentX;
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
     * Slide identifier
     *
     * @var string
     */
    protected $identifier;
    /**
     * Hash index
     *
     * @var string
     */
    protected $hashIndex;
    /**
     * Parent presentation
     *
     * @var PhpPresentation
     */
    protected $parent;
    /**
     * Background of the slide
     *
     * @var AbstractBackground
     */
    protected $background;

    /**
     * Get collection of shapes
     *
     * @return \ArrayObject|\PhpOffice\PhpPresentation\AbstractShape[]
     */
    public function getShapeCollection()
    {
        return $this->shapeCollection;
    }

    /**
     * Get collection of shapes
     *
     * @return AbstractSlide
     */
    public function setShapeCollection($shapeCollection = array())
    {
        $this->shapeCollection = $shapeCollection;
        return $this;
    }

    /**
     * Add shape to slide
     *
     * @param  \PhpOffice\PhpPresentation\AbstractShape $shape
     * @return \PhpOffice\PhpPresentation\AbstractShape
     */
    public function addShape(AbstractShape $shape)
    {
        $shape->setContainer($this);
        return $shape;
    }

    /**
     * Get X Offset
     *
     * @return int
     */
    public function getOffsetX()
    {
        if ($this->offsetX === null) {
            $offsets = GeometryCalculator::calculateOffsets($this);
            $this->offsetX = $offsets[GeometryCalculator::X];
            $this->offsetY = $offsets[GeometryCalculator::Y];
        }
        return $this->offsetX;
    }

    /**
     * Get Y Offset
     *
     * @return int
     */
    public function getOffsetY()
    {
        if ($this->offsetY === null) {
            $offsets = GeometryCalculator::calculateOffsets($this);
            $this->offsetX = $offsets[GeometryCalculator::X];
            $this->offsetY = $offsets[GeometryCalculator::Y];
        }
        return $this->offsetY;
    }

    /**
     * Get X Extent
     *
     * @return int
     */
    public function getExtentX()
    {
        if ($this->extentX === null) {
            $extents = GeometryCalculator::calculateExtents($this);
            $this->extentX = $extents[GeometryCalculator::X];
            $this->extentY = $extents[GeometryCalculator::Y];
        }
        return $this->extentX;
    }

    /**
     * Get Y Extent
     *
     * @return int
     */
    public function getExtentY()
    {
        if ($this->extentY === null) {
            $extents = GeometryCalculator::calculateExtents($this);
            $this->extentX = $extents[GeometryCalculator::X];
            $this->extentY = $extents[GeometryCalculator::Y];
        }
        return $this->extentY;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->identifier . __CLASS__);
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
     * Create rich text shape
     *
     * @return \PhpOffice\PhpPresentation\Shape\RichText
     */
    public function createRichTextShape()
    {
        $shape = new RichText();
        $this->addShape($shape);
        return $shape;
    }

    /**
     * Create line shape
     *
     * @param  int $fromX Starting point x offset
     * @param  int $fromY Starting point y offset
     * @param  int $toX Ending point x offset
     * @param  int $toY Ending point y offset
     * @return \PhpOffice\PhpPresentation\Shape\Line
     */
    public function createLineShape($fromX, $fromY, $toX, $toY)
    {
        $shape = new Line($fromX, $fromY, $toX, $toY);
        $this->addShape($shape);
        return $shape;
    }

    /**
     * Create chart shape
     *
     * @return \PhpOffice\PhpPresentation\Shape\Chart
     */
    public function createChartShape()
    {
        $shape = new Chart();
        $this->addShape($shape);
        return $shape;
    }

    /**
     * Create drawing shape
     *
     * @return \PhpOffice\PhpPresentation\Shape\Drawing\File
     */
    public function createDrawingShape()
    {
        $shape = new File();
        $this->addShape($shape);
        return $shape;
    }

    /**
     * Create table shape
     *
     * @param  int $columns Number of columns
     * @return \PhpOffice\PhpPresentation\Shape\Table
     */
    public function createTableShape($columns = 1)
    {
        $shape = new Table($columns);
        $this->addShape($shape);
        return $shape;
    }

    /**
     * Creates a group within this slide
     *
     * @return \PhpOffice\PhpPresentation\Shape\Group
     */
    public function createGroup()
    {
        $shape = new Group();
        $this->addShape($shape);
        return $shape;
    }

    /**
     * Get parent
     *
     * @return PhpPresentation
     */
    public function getParent()
    {
        return $this->parent;
    }

    /**
     * Re-bind parent
     *
     * @param  \PhpOffice\PhpPresentation\PhpPresentation $parent
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function rebindParent(PhpPresentation $parent)
    {
        $this->parent->removeSlideByIndex($this->parent->getIndex($this));
        $this->parent = $parent;
        return $this;
    }

    /**
     * @return AbstractBackground
     */
    public function getBackground()
    {
        return $this->background;
    }

    /**
     * @param AbstractBackground $background
     * @return Slide
     */
    public function setBackground(AbstractBackground $background = null)
    {
        $this->background = $background;
        return $this;
    }

    /**
     *
     * @return \PhpOffice\PhpPresentation\Slide\Transition
     */
    public function getTransition()
    {
        return $this->slideTransition;
    }

    /**
     *
     * @param \PhpOffice\PhpPresentation\Slide\Transition $transition
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function setTransition(Transition $transition = null)
    {
        $this->slideTransition = $transition;
        return $this;
    }

    /**
     * @return string
     */
    public function getRelsIndex()
    {
        return $this->relsIndex;
    }

    /**
     * @param string $indexName
     */
    public function setRelsIndex($indexName)
    {
        $this->relsIndex = $indexName;
    }
}
