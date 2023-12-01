<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Writer\PDF\PDFWriterInterface;

abstract class AbstractWriter
{
    /**
     * Private unique hash table.
     *
     * @var HashTable
     */
    protected $oDrawingHashTable;

    /**
     * Private PhpPresentation.
     *
     * @var null|PhpPresentation
     */
    protected $oPresentation;

    /**
     * @var null|PDFWriterInterface
     */
    protected $pdfAdapter;

    /**
     * @var null|ZipInterface
     */
    protected $oZipAdapter;

    /**
     * Get drawing hash table.
     */
    public function getDrawingHashTable(): HashTable
    {
        return $this->oDrawingHashTable;
    }

    public function getPDFAdapter(): ?PDFWriterInterface
    {
        return $this->pdfAdapter;
    }

    /**
     * Get PhpPresentation object.
     */
    public function getPhpPresentation(): ?PhpPresentation
    {
        return $this->oPresentation;
    }

    public function setPDFAdapter(PDFWriterInterface $pdfAdapter): self
    {
        $this->pdfAdapter = $pdfAdapter;

        return $this;
    }

    /**
     * Get PhpPresentation object.
     *
     * @param null|PhpPresentation $pPhpPresentation PhpPresentation object
     *
     * @return self
     */
    public function setPhpPresentation(?PhpPresentation $pPhpPresentation = null)
    {
        $this->oPresentation = $pPhpPresentation;

        return $this;
    }

    public function setZipAdapter(ZipInterface $oZipAdapter): self
    {
        $this->oZipAdapter = $oZipAdapter;

        return $this;
    }

    public function getZipAdapter(): ?ZipInterface
    {
        return $this->oZipAdapter;
    }

    /**
     * Get an array of all drawings.
     *
     * @return array<int, AbstractShape>
     */
    protected function allDrawings(): array
    {
        // Get an array of all drawings
        $aDrawings = [];

        // Get an array of all master slides
        $aSlideMasters = $this->getPhpPresentation()->getAllMasterSlides();

        $aSlideMasterLayouts = array_map(function ($oSlideMaster) {
            return $oSlideMaster->getAllSlideLayouts();
        }, $aSlideMasters);

        // Get an array of all slide layouts
        $aSlideLayouts = [];
        array_walk_recursive($aSlideMasterLayouts, function ($oSlideLayout) use (&$aSlideLayouts): void {
            $aSlideLayouts[] = $oSlideLayout;
        });

        // Loop through PhpPresentation
        foreach (array_merge($this->getPhpPresentation()->getAllSlides(), $aSlideMasters, $aSlideLayouts) as $oSlide) {
            $arrayReturn = $this->iterateCollection($oSlide->getShapeCollection());
            $aDrawings = array_merge($aDrawings, $arrayReturn);
        }

        return $aDrawings;
    }

    /**
     * @param array<int, AbstractShape> $collection
     *
     * @return array<int, AbstractShape>
     */
    private function iterateCollection(array $collection): array
    {
        $arrayReturn = [];

        foreach ($collection as $oShape) {
            if ($oShape instanceof AbstractDrawingAdapter) {
                $arrayReturn[] = $oShape;
            } elseif ($oShape instanceof Chart) {
                $arrayReturn[] = $oShape;
            } elseif ($oShape instanceof Group) {
                $arrayGroup = $this->iterateCollection($oShape->getShapeCollection());
                $arrayReturn = array_merge($arrayReturn, $arrayGroup);
            }
        }

        return $arrayReturn;
    }
}
