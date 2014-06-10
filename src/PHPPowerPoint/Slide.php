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

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\Chart;
use PhpOffice\PhpPowerpoint\Shape\Drawing;
use PhpOffice\PhpPowerpoint\Shape\Line;
use PhpOffice\PhpPowerpoint\Shape\RichText;
use PhpOffice\PhpPowerpoint\Shape\Table;
use PhpOffice\PhpPowerpoint\Slide\Layout;

/**
 * PHPPowerPoint_Slide
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Slide
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Slide implements IComparable
{
    /**
     * Parent presentation
     *
     * @var PHPPowerPoint
     */
    private $_parent;

    /**
     * Collection of shapes
     *
     * @var PHPPowerPoint_Shape[]
     */
    private $_shapeCollection = null;

    /**
     * Slide identifier
     *
     * @var string
     */
    private $_identifier;

    /**
     * Slide layout
     *
     * @var string
     */
    private $_slideLayout;

    /**
     * Slide master id
     *
     * @var string
     */
    private $_slideMasterId = 1;

    /**
     * Create a new slide
     *
     * @param PHPPowerPoint $pParent
     */
    public function __construct(PhpPowerpoint $pParent = null)
    {
        // Set parent
        $this->_parent = $pParent;
        
        $this->_slideLayout = Slide\Layout::BLANK;

        // Shape collection
        $this->_shapeCollection = new \ArrayObject();

        // Set identifier
        $this->_identifier = md5(rand(0, 9999) . time());
    }

    /**
     * Get collection of shapes
     *
     * @return PHPPowerPoint_Shape[]
     */
    public function getShapeCollection()
    {
        return $this->_shapeCollection;
    }

    /**
     * Add shape to slide
     *
     * @param  PHPPowerPoint_Shape $shape
     * @return PHPPowerPoint_Shape
     */
    public function addShape(Shape $shape)
    {
        $shape->setSlide($this);

        return $shape;
    }

    /**
     * Create rich text shape
     *
     * @return \PHPPowerPoint\Shape\RichText
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
     * @return PHPPowerPoint_Shape_Line
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
     * @return PHPPowerPoint_Shape_Chart
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
     * @return PHPPowerPoint_Shape_Drawing
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
     * @return PHPPowerPoint_Shape_Table
     */
    public function createTableShape($columns = 1)
    {
        $shape = new Table($columns);
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
        return $this->_parent;
    }

    /**
     * Re-bind parent
     *
     * @param  PHPPowerPoint       $parent
     * @return PHPPowerPoint_Slide
     */
    public function rebindParent(PHPPowerPoint $parent)
    {
        $this->_parent->removeSlideByIndex($this->_parent->getIndex($this));
        $this->_parent = $parent;

        return $this;
    }

    /**
     * Get slide layout
     *
     * @return string
     */
    public function getSlideLayout()
    {
        return $this->_slideLayout;
    }

    /**
     * Set slide layout
     *
     * @param  string              $layout
     * @return PHPPowerPoint_Slide
     */
    public function setSlideLayout($layout = Layout::BLANK)
    {
        $this->_slideLayout = $layout;

        return $this;
    }

    /**
     * Get slide master id
     *
     * @return int
     */
    public function getSlideMasterId()
    {
        return $this->_slideMasterId;
    }

    /**
     * Set slide master id
     *
     * @param  int                 $masterId
     * @return PHPPowerPoint_Slide
     */
    public function setSlideMasterId($masterId = 1)
    {
        $this->_slideMasterId = $masterId;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->_identifier . __CLASS__);
    }

    /**
     * Hash index
     *
     * @var string
     */
    private $_hashIndex;

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
        return $this->_hashIndex;
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
        $this->_hashIndex = $value;
    }

    /**
     * Copy slide (!= clone!)
     *
     * @return PHPPowerPoint_Slide
     */
    public function copy()
    {
        $copied = clone $this;

        return $copied;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        foreach ($this as $key => $val) {
            if (is_object($val) || (is_array($val))) {
                $this->{$key} = unserialize(serialize($val));
            }
        }
    }
}
