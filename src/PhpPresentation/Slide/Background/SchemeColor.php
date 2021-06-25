<?php

namespace PhpOffice\PhpPresentation\Slide\Background;

use PhpOffice\PhpPresentation\Slide\AbstractBackground;
use PhpOffice\PhpPresentation\Style\SchemeColor as StyleSchemeColor;

class SchemeColor extends AbstractBackground
{
    /**
     * @var StyleSchemeColor|null
     */
    protected $schemeColor;

    /**
     * @return $this
     */
    public function setSchemeColor(StyleSchemeColor $color = null): self
    {
        $this->schemeColor = $color;

        return $this;
    }

    public function getSchemeColor(): ?StyleSchemeColor
    {
        return $this->schemeColor;
    }
}
