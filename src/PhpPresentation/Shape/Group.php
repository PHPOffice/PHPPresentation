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

namespace PhpOffice\PhpPresentation\Shape;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Traits\ShapeCollection;

class Group extends AbstractShape implements ShapeContainerInterface
{
    use ShapeCollection;

    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Get X Offset.
     */
    public function getOffsetX(): int
    {
        $offsetX = null;

        foreach ($this->getShapeCollection() as $shape) {
            if ($offsetX === null) {
                $offsetX = $shape->getOffsetX();
            } else {
                $offsetX = \min($offsetX, $shape->getOffsetX());
            }
        }

        return $offsetX ?? 0;
    }

    /**
     * Change the X offset by moving all contained shapes.
     *
     * @return $this
     */
    public function setOffsetX(int $pValue = 0): self
    {
        $offsetX = $this->getOffsetX();
        $diff = $pValue - $offsetX;

        foreach ($this->getShapeCollection() as $shape) {
            $shape->setOffsetX($shape->getOffsetX() + $diff);
        }

        return $this;
    }

    /**
     * Get Y Offset.
     */
    public function getOffsetY(): int
    {
        $offsetY = null;

        foreach ($this->getShapeCollection() as $shape) {
            if ($offsetY === null) {
                $offsetY = $shape->getOffsetY();
            } else {
                $offsetY = \min($offsetY, $shape->getOffsetY());
            }
        }

        return $offsetY ?? 0;
    }

    /**
     * Change the Y offset by moving all contained shapes.
     *
     * @return $this
     */
    public function setOffsetY(int $pValue = 0): self
    {
        $offsetY = $this->getOffsetY();
        $diff = $pValue - $offsetY;

        foreach ($this->getShapeCollection() as $shape) {
            $shape->setOffsetY($shape->getOffsetY() + $diff);
        }

        return $this;
    }

    /**
     * Get X Extent.
     */
    public function getExtentX(): int
    {
        $extentX = 0;

        foreach ($this->getShapeCollection() as $shape) {
            $extentX = \max($extentX, $shape->getOffsetX() + $shape->getWidth());
        }

        return $extentX - $this->getOffsetX();
    }

    /**
     * Get Y Extent.
     */
    public function getExtentY(): int
    {
        $extentY = 0;

        foreach ($this->getShapeCollection() as $shape) {
            $extentY = \max($extentY, $shape->getOffsetY() + $shape->getHeight());
        }

        return $extentY - $this->getOffsetY();
    }

    /**
     * Calculate the width based on the size/position of the contained shapes.
     */
    public function getWidth(): int
    {
        return $this->getExtentX();
    }

    /**
     * Calculate the height based on the size/position of the contained shapes.
     */
    public function getHeight(): int
    {
        return $this->getExtentY();
    }

    /**
     * Ignores setting the width, preserving the default behavior.
     *
     * @return $this
     */
    public function setWidth(int $pValue = 0): self
    {
        return $this;
    }

    /**
     * Ignores setting the height, preserving the default behavior.
     *
     * @return $this
     */
    public function setHeight(int $pValue = 0): self
    {
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
     */
    public function createTableShape(int $columns = 1): Table
    {
        $shape = new Table($columns);
        $this->addShape($shape);

        return $shape;
    }
}
