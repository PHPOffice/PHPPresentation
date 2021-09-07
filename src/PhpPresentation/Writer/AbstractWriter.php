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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer;

use ArrayIterator;
use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Group;

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
     * @var PhpPresentation|null
     */
    protected $oPresentation;

    /**
     * @var ZipInterface|null
     */
    protected $oZipAdapter;

    /**
     * Get drawing hash table.
     */
    public function getDrawingHashTable(): HashTable
    {
        return $this->oDrawingHashTable;
    }

    /**
     * Get PhpPresentation object.
     */
    public function getPhpPresentation(): ?PhpPresentation
    {
        return $this->oPresentation;
    }

    /**
     * Get PhpPresentation object.
     *
     * @param PhpPresentation|null $pPhpPresentation PhpPresentation object
     *
     * @return self
     */
    public function setPhpPresentation(PhpPresentation $pPhpPresentation = null)
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
        array_walk_recursive($aSlideMasterLayouts, function ($oSlideLayout) use (&$aSlideLayouts) {
            $aSlideLayouts[] = $oSlideLayout;
        });

        // Loop through PhpPresentation
        foreach (array_merge($this->getPhpPresentation()->getAllSlides(), $aSlideMasters, $aSlideLayouts) as $oSlide) {
            $arrayReturn = $this->iterateCollection($oSlide->getShapeCollection()->getIterator());
            $aDrawings = array_merge($aDrawings, $arrayReturn);
        }

        return $aDrawings;
    }

    /**
     * @param ArrayIterator<int, AbstractShape> $oIterator
     *
     * @return array<int, AbstractShape>
     */
    private function iterateCollection(ArrayIterator $oIterator): array
    {
        $arrayReturn = [];
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
