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

namespace PhpOffice\PhpPowerPoint\Shape;

use PhpOffice\PhpPowerpoint\AbstractShape;
use PhpOffice\PhpPowerpoint\GeometryCalculator;
use PHPOffice\PhpPowerPoint\ShapeContainerInterface;
use PhpOffice\PhpPowerpoint\Shape\Drawing;
use PhpOffice\PhpPowerpoint\Shape\Line;
use PhpOffice\PhpPowerpoint\Shape\RichText;
use PhpOffice\PhpPowerpoint\Shape\Table;

class Group extends AbstractShape implements ShapeContainerInterface
{
  /**
   * Collection of shapes
   *
   * @var \ArrayObject|\PhpOffice\PhpPowerpoint\AbstractShape[]
   */
  private $shapeCollection = null;

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
  
  public function __construct() {
    parent::__construct();
    
    // For logic purposes.
    $this->offsetX = null;
    $this->offsetY = null;
    
    // Shape collection
    $this->shapeCollection = new \ArrayObject(); 
  }
  
  /**
   * Get collection of shapes
   *
   * @return \PhpOffice\PhpPowerpoint\AbstractShape[]
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
   * Get X Offset
   *
   * @return int
   */
  public function getOffsetX()
  {
      if ($this->offsetX === null)
      {
          $offsets = GeometryCalculator::calculateOffsets($this);
          $this->offsetX = $offsets[GeometryCalculator::X];
          $this->offsetY = $offsets[GeometryCalculator::Y];
      }

      return $this->offsetX;
  }

  /**
   * Ignores setting the X Offset, preserving the default behavior.
   *
   * @param  int                 $pValue
   * @return self
   */
  public function setOffsetX($pValue = 0)
  {
      return $this;
  }

  /**
   * Get Y Offset
   *
   * @return int
   */
  public function getOffsetY()
  {
      if ($this->offsetY === null)
      {
          $offsets = GeometryCalculator::calculateOffsets($this);
          $this->offsetX = $offsets[GeometryCalculator::X];
          $this->offsetY = $offsets[GeometryCalculator::Y];
      }
      return $this->offsetY;
  }

  /**
   * Ignores setting the Y Offset, preserving the default behavior.
   *
   * @param  int                 $pValue
   * @return self
   */
  public function setOffsetY($pValue = 0)
  {
      return $this;
  }

  /**
   * Get X Extent
   *
   * @return int
   */
  public function getExtentX()
  {
      if ($this->extentX === null)
      {
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
      if ($this->extentY === null)
      {
          $extents = GeometryCalculator::calculateExtents($this);
          $this->extentX = $extents[GeometryCalculator::X];
          $this->extentY = $extents[GeometryCalculator::Y];
      }
      return $this->extentY;
  }

  /**
   * Ignores setting the width, preserving the default behavior.
   *
   * @param  int                 $pValue
   * @return self
   */
  public function setWidth($pValue = 0)
  {
      return $this;
  }

  /**
   * Ignores setting the height, preserving the default behavior.
   *
   * @param  int                 $pValue
   * @return self
   */
  public function setHeight($pValue = 0)
  {
      return $this;
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
}