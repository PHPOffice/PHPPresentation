<?php

namespace PhpOffice\PhpPresentation\Shape\Drawing;

use PhpOffice\PhpPresentation\Shape\AbstractGraphic;

abstract class AbstractDrawingAdapter extends AbstractGraphic
{
    abstract public function getContents(): string;

    abstract public function getExtension(): string;

    abstract public function getIndexedFilename(): string;

    abstract public function getMimeType(): string;

    abstract public function getPath(): string;

    /**
     * @return self
     */
    abstract public function setPath(string $path);
}
