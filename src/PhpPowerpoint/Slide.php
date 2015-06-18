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

use PhpOffice\PhpPowerpoint\Shape\Chart;
use PhpOffice\PhpPowerpoint\Shape\Drawing;
use PhpOffice\PhpPowerpoint\Shape\Line;
use PhpOffice\PhpPowerpoint\Shape\RichText;
use PhpOffice\PhpPowerpoint\Shape\Table;
use PhpOffice\PhpPowerpoint\Slide\Layout;
use PhpOffice\PhpPowerpoint\Slide\Note;
use PhpOffice\PhpPowerpoint\Slide\Transition;

/**
 * Slide class
 */
class Slide implements ComparableInterface, ShapeContainerInterface
{
    /**
     * Parent presentation
     *
     * @var PHPPowerPoint
     */
    private $parent;

    /**
     * Collection of shapes
     *
     * @var \ArrayObject|\PhpOffice\PhpPowerpoint\AbstractShape[]
     */
    private $shapeCollection = null;

    /**
     * Slide identifier
     *
     * @var string
     */
    private $identifier;

    /**
     * Slide layout
     *
     * @var string
     */
    private $slideLayout;

    /**
     * Slide master id
     *
     * @var integer
     */
    private $slideMasterId = 1;

    /**
     *
     * @var \PhpOffice\PhpPowerpoint\Slide\Note
     */
    private $slideNote;
    /**
     *
     * @var \PhpOffice\PhpPowerpoint\Slide\Transition
     */
    private $slideTransition;
    
    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

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
     * Extent X
     *
     * @var int
     */
    protected $extentX;

    /**
     * Extent Y
     *
     * @var int
     */
    protected $extentY;

    /**
     * Create a new slide
     *
     * @param PHPPowerPoint $pParent
     */
    public function __construct(PhpPowerpoint $pParent = null)
    {
        // Set parent
        $this->parent = $pParent;

        $this->slideLayout = Slide\Layout::BLANK;

        // Shape collection
        $this->shapeCollection = new \ArrayObject();

        // Set identifier
        $this->identifier = md5(rand(0, 9999) . time());
    }

    /**
     * Get collection of shapes
     *
     * @return \ArrayObject|\PhpOffice\PhpPowerpoint\AbstractShape[]
     */
    public function getShapeCollection()
    {
        return $this->shapeCollection;
    }

    /**
     * Add shape to slide
     *
     * @param  \PhpOffice\PhpPowerpoint\AbstractShape $shape
     * @return \PhpOffice\PhpPowerpoint\AbstractShape
     */
    public function addShape(AbstractShape $shape)
    {
        $shape->setContainer($this);

        return $shape;
    }

    /**
     * Create rich text shape
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText
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
     * @param  int                      $fromX Starting point x offset
     * @param  int                      $fromY Starting point y offset
     * @param  int                      $toX   Ending point x offset
     * @param  int                      $toY   Ending point y offset
     * @return \PhpOffice\PhpPowerpoint\Shape\Line
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Drawing
     */
    public function createDrawingShape()
    {
        $shape = new Drawing();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create table shape
     *
     * @param  int                       $columns Number of columns
     * @return \PhpOffice\PhpPowerpoint\Shape\Table
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
     * @return \PhpOffice\PhpPowerpoint\Shape\Group
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
     * @return PHPPowerPoint
     */
    public function getParent()
    {
        return $this->parent;
    }

    /**
     * Re-bind parent
     *
     * @param  \PhpOffice\PhpPowerpoint\PhpPowerpoint       $parent
     * @return \PhpOffice\PhpPowerpoint\Slide
     */
    public function rebindParent(PhpPowerpoint $parent)
    {
        $this->parent->removeSlideByIndex($this->parent->getIndex($this));
        $this->parent = $parent;

        return $this;
    }

    /**
     * Get slide layout
     *
     * @return string
     */
    public function getSlideLayout()
    {
        return $this->slideLayout;
    }

    /**
     * Set slide layout
     *
     * @param  string              $layout
     * @return \PhpOffice\PhpPowerpoint\Slide
     */
    public function setSlideLayout($layout = Layout::BLANK)
    {
        $this->slideLayout = $layout;

        return $this;
    }

    /**
     * Get slide master id
     *
     * @return int
     */
    public function getSlideMasterId()
    {
        return $this->slideMasterId;
    }

    /**
     * Set slide master id
     *
     * @param  int                 $masterId
     * @return \PhpOffice\PhpPowerpoint\Slide
     */
    public function setSlideMasterId($masterId = 1)
    {
        $this->slideMasterId = $masterId;

        return $this;
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
     * Copy slide (!= clone!)
     *
     * @return \PhpOffice\PhpPowerpoint\Slide
     */
    public function copy()
    {
        $copied = clone $this;

        return $copied;
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
     *
     * @return \PhpOffice\PhpPowerpoint\Slide\Note
     */
    public function getNote()
    {
        if (is_null($this->slideNote)) {
            $this->setNote();
        }
        return $this->slideNote;
    }

    /**
     *
     * @param \PhpOffice\PhpPowerpoint\Slide\Note $note
     * @return \PhpOffice\PhpPowerpoint\Slide
     */
    public function setNote(Note $note = null)
    {
        $this->slideNote = (is_null($note) ? new Note() : $note);
        $this->slideNote->setParent($this);

        return $this;
    }

    /**
     *
     * @return \PhpOffice\PhpPowerpoint\Slide\Transition
     */
    public function getTransition()
    {
        return $this->slideTransition;
    }

    /**
     *
     * @param \PhpOffice\PhpPowerpoint\Slide\Transition $transition
     * @return \PhpOffice\PhpPowerpoint\Slide
     */
    public function setTransition(Transition $transition = null)
    {
        $this->slideTransition = $transition;

        return $this;
    }
}
