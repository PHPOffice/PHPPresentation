<?php
namespace PhpOffice\PhpPresentation\Slide;

use PhpOffice\PhpPresentation\AbstractShape;

class Animation
{
    /**
     * @var array
     */
    protected $shapeCollection = array();

    /**
     * @param AbstractShape $shape
     * @return Animation
     */
    public function addShape(AbstractShape $shape)
    {
        $this->shapeCollection[] = $shape;
        return $this;
    }

    /**
     * @return \PhpOffice\PhpPresentation\AbstractShape[]
     */
    public function getShapeCollection()
    {
        return $this->shapeCollection;
    }

    /**
     * @param \PhpOffice\PhpPresentation\AbstractShape[] $array
     * @return Animation
     */
    public function setShapeCollection(array $array = array())
    {
        $this->shapeCollection = $array;
        return $this;
    }
}
