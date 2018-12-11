<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

class Mimetype extends AbstractDecoratorWriter
{
    /**
     * @return ZipInterface
     */
    public function render()
    {
        $this->getZip()->addFromString('mimetype', 'application/vnd.oasis.opendocument.presentation');
        return $this->getZip();
    }
}
