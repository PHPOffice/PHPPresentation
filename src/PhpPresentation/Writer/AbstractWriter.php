<?php

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\PhpPresentation;

abstract class AbstractWriter
{
    /**
     * Private unique hash table
     *
     * @var \PhpOffice\PhpPresentation\HashTable
     */
    protected $oDrawingHashTable;

    /**
     * Private PhpPresentation
     *
     * @var PhpPresentation
     */
    protected $oPresentation;

    /**
     * @var ZipInterface
     */
    protected $oZipAdapter;

    /**
     * Get drawing hash table
     *
     * @return \PhpOffice\PhpPresentation\HashTable
     */
    public function getDrawingHashTable()
    {
        return $this->oDrawingHashTable;
    }

    /**
     * Get PhpPresentation object
     *
     * @return PhpPresentation
     * @throws \Exception
     */
    public function getPhpPresentation()
    {
        if (empty($this->oPresentation)) {
            throw new \Exception("No PhpPresentation assigned.");
        }
        return $this->oPresentation;
    }

    /**
     * Get PhpPresentation object
     *
     * @param  PhpPresentation                       $pPhpPresentation PhpPresentation object
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Writer\ODPresentation
     */
    public function setPhpPresentation(PhpPresentation $pPhpPresentation = null)
    {
        $this->oPresentation = $pPhpPresentation;
        return $this;
    }


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
