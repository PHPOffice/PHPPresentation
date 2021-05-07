<?php
/**
 * @author Muhammad Hasan Shahid <m.hasan.shahid12@gmail.com>
 * Mindline Analytics Gmbh
 */

namespace PhpOffice\PhpPresentation\Shape;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Style\Border;


/**
 * Class Triangle
 * @package PhpOffice\PhpPresentation\Shape
 * Make shape triangle
 */
class Triangle extends AbstractShape implements ComparableInterface
{
    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Triangle instance
     *
     * @param int $fromX
     * @param int $fromY
     * @param int $toX
     * @param int $toY
     * @param int $rotation used for rotation clockwise or anti-clockwise
     */
    public function __construct($fromX, $fromY, $toX, $toY, $rotation)
    {
        parent::__construct();
        $this->getBorder()->setLineStyle(Border::LINE_SINGLE);

        $this->setOffsetX($fromX);
        $this->setOffsetY($fromY);
        $this->setWidth($toX - $fromX);
        $this->setHeight($toY - $fromY);
        $this->setRotation($rotation);
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->getBorder()->getLineStyle() . parent::getHashCode() . __CLASS__);
    }

}
