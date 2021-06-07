<?php

namespace PhpOffice\PhpPresentation\Shape\Drawing;

use PhpOffice\PhpPresentation\Shape\AbstractGraphic;

abstract class AbstractDrawingAdapter extends AbstractGraphic
{
    /**
     * @return string
     */
    abstract public function getContents(): string;

    /**
     * @return string
     */
    abstract public function getExtension(): string;

    /**
     * @return string
     */
    abstract public function getIndexedFilename(): string;

    /**
     * @return string
     */
    abstract public function getMimeType(): string;

    /**
     * @return string
     */
    abstract public function getPath(): string;

    /**
     * @param string $path
     * @return self
     */
    abstract public function setPath(string $path);
}
