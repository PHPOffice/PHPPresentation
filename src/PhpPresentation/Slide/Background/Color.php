<?php

namespace PhpOffice\PhpPresentation\Slide\Background;

use PhpOffice\PhpPresentation\Slide\AbstractBackground;
use PhpOffice\PhpPresentation\Style\Color as StyleColor;

class Color extends AbstractBackground
{
    /**
     * @var StyleColor
     */
    protected $color;

    /**
     * @param StyleColor|null $color
     * @return $this
     */
    public function setColor(StyleColor $color = null)
    {
        $this->color = $color;
        return $this;
    }

    /**
     * @return StyleColor
     */
    public function getColor()
    {
        return $this->color;
    }
}
