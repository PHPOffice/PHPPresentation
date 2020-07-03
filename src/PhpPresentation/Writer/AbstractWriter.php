<?php

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Group;

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

    /**
     * Get an array of all drawings
     *
     * @return \PhpOffice\PhpPresentation\Shape\AbstractDrawing[] All drawings in PhpPresentation
     * @throws \Exception
     */
    protected function allDrawings()
    {
        // Get an array of all drawings
        $aDrawings  = array();

        // Loop through PhpPresentation
        foreach (array_merge($this->getPhpPresentation()->getAllSlides(), $this->getPhpPresentation()->getAllMasterSlides()) as $oSlide) {
            $arrayReturn = $this->iterateCollection($oSlide->getShapeCollection()->getIterator());
            $aDrawings = array_merge($aDrawings, $arrayReturn);
        }

        return $aDrawings;
    }

    private function iterateCollection(\ArrayIterator $oIterator)
    {
        $arrayReturn = array();
        if ($oIterator->count() <= 0) {
            return $arrayReturn;
        }

        while ($oIterator->valid()) {
            $oShape = $oIterator->current();
            if ($oShape instanceof AbstractDrawingAdapter) {
                $arrayReturn[] = $oShape;
            } elseif ($oShape instanceof Chart) {
                $arrayReturn[] = $oShape;
            } elseif ($oShape instanceof Group) {
                $arrayGroup = $this->iterateCollection($oShape->getShapeCollection()->getIterator());
                $arrayReturn = array_merge($arrayReturn, $arrayGroup);
            }
            $oIterator->next();
        }
        return $arrayReturn;
    }
}
