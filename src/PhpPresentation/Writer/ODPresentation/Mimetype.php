<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

class Mimetype extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     *
     * @throws \Exception
     */
    public function render()
    {
        $this->getZip()->addFromString('mimetype', 'application/vnd.oasis.opendocument.presentation');

        return $this->getZip();
    }
}
