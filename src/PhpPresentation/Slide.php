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

use PhpOffice\PhpPresentation\GeometryCalculator;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide\AbstractBackground;
use PhpOffice\PhpPresentation\Slide\Layout;
use PhpOffice\PhpPresentation\Slide\Note;
use PhpOffice\PhpPresentation\Slide\Transition;

/**
 * Slide class
 */
class Slide implements ComparableInterface, ShapeContainerInterface
{
    /**
     * The slide is shown in presentation
     * @var bool
     */
    protected $isVisible = true;

    /**
     * Parent presentation
     *
     * @var PhpPresentation
     */
    private $parent;

    /**
     * Collection of shapes
     *
     * @var \ArrayObject|\PhpOffice\PhpPresentation\AbstractShape[]
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
     * @var \PhpOffice\PhpPresentation\Slide\Note
     */
    private $slideNote;

    /**
     *
     * @var \PhpOffice\PhpPresentation\Slide\Transition
     */
    private $slideTransition;
  
    /**
     *
     * @var \PhpOffice\PhpPresentation\Slide\Animation[]
     */
    protected $animations = array();
    
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
     * Name of the title
     *
     * @var string
     */
    protected $name;

    /**
     * Background of the slide
     *
     * @var AbstractBackground
     */
    protected $background;

    /**
     * Create a new slide
     *
     * @param PhpPresentation $pParent
     */
    public function __construct(PhpPresentation $pParent = null)
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
     * @return \ArrayObject|\PhpOffice\PhpPresentation\AbstractShape[]
     */
    public function getShapeCollection()
    {
        return $this->shapeCollection;
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
     * @param  int                      $fromX Starting point x offset
     * @param  int                      $fromY Starting point y offset
     * @param  int                      $toX   Ending point x offset
     * @param  int                      $toY   Ending point y offset
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
        $shape = new Drawing\File();
        $this->addShape($shape);

        return $shape;
    }

    /**
     * Create table shape
     *
     * @param  int                       $columns Number of columns
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
     * @param  \PhpOffice\PhpPresentation\PhpPresentation       $parent
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function rebindParent(PhpPresentation $parent)
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
     * @return \PhpOffice\PhpPresentation\Slide
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
     * @return \PhpOffice\PhpPresentation\Slide
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
     * @return \PhpOffice\PhpPresentation\Slide
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
     * @return \PhpOffice\PhpPresentation\Slide\Note
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
     * @param \PhpOffice\PhpPresentation\Slide\Note $note
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function setNote(Note $note = null)
    {
        $this->slideNote = (is_null($note) ? new Note() : $note);
        $this->slideNote->setParent($this);

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
     * Get the name of the slide
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Set the name of the slide
     * @param string $name
     * @return $this
     */
    public function setName($name = null)
    {
        $this->name = $name;
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
     * @return $this
     */
    public function setBackground(AbstractBackground $background = null)
    {
        $this->background = $background;
        return $this;
    }

    /**
     * @return boolean
     */
    public function isVisible()
    {
        return $this->isVisible;
    }

    /**
     * @param boolean $value
     * @return Slide
     */
    public function setIsVisible($value = true)
    {
        $this->isVisible = (bool)$value;
        return $this;
    }

    /**
     * Add an animation to the slide
     *
     * @param  \PhpOffice\PhpPresentation\Slide\Animation
     * @return Slide
     */
    public function addAnimation($animation)
    {
        $this->animations[] = $animation;
        return $this;
    }

    /**
     * Get collection of animations
     *
     * @return \PhpOffice\PhpPresentation\Slide\Animation
     */
    public function getAnimations()
    {
        return $this->animations;
    }

    /**
     * Set collection of animations
     * @param \PhpOffice\PhpPresentation\Slide\Animation[] $array
     * @return Slide
     */
    public function setAnimations(array $array = array())
    {
        $this->animations = $array;
        return $this;
    }
}
