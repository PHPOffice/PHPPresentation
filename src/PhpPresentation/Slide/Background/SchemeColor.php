<?php

namespace PhpOffice\PhpPresentation\Slide\Background;

use PhpOffice\PhpPresentation\Slide\AbstractBackground;
use PhpOffice\PhpPresentation\Style\SchemeColor as StyleSchemeColor;

class SchemeColor extends AbstractBackground
{
    /**
     * @var StyleSchemeColor
     */
    protected $schemeColor;

    /**
     * @param StyleSchemeColor|null $color
     * @return $this
     */
    public function setSchemeColor(StyleSchemeColor $color = null)
    {
        $this->schemeColor = $color;
        return $this;
    }

    /**
     * @return StyleSchemeColor
     */
    public function getSchemeColor()
    {
        return $this->schemeColor;
    }
}
