<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;

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
