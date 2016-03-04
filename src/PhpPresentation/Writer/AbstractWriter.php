<?php

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\Common\Adapter\Zip\ZipInterface;

abstract class AbstractWriter
{
    /**
     * @var ZipInterface
     */
    protected $oZipAdapter;

    /**
     * @param ZipInterface $oZipAdapter
     * @return $this
     */
    public function setZipAdapter(ZipInterface $oZipAdapter)
    {
        $this->oZipAdapter = $oZipAdapter;
        return $this;
    }

    /**
     * @return ZipInterface
     */
    public function getZipAdapter()
    {
        return $this->oZipAdapter;
    }
}
