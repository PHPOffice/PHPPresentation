<?php

namespace PhpOffice\PhpPresentation\Slide;

use PhpOffice\PhpPresentation\AbstractShape;

class Animation
{
    /**
     * @var array<AbstractShape>
     */
    protected $shapeCollection = [];

    /**
     * @return Animation
     */
    public function addShape(AbstractShape $shape)
    {
        $this->shapeCollection[] = $shape;

        return $this;
    }

    /**
     * @return array<AbstractShape>
     */
    public function getShapeCollection(): array
    {
        return $this->shapeCollection;
    }

    /**
     * @param array<AbstractShape> $array
     *
     * @return Animation
     */
    public function setShapeCollection(array $array = [])
    {
        $this->shapeCollection = $array;

        return $this;
    }
}
