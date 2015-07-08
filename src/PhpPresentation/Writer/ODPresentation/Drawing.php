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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AbstractDrawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Table;

/**
 * \PhpOffice\PhpPresentation\Writer\ODPresentation\Drawing
 */
class Drawing extends AbstractPart
{
    /**
     * Get an array of all drawings
     *
     * @param  PhpPresentation                 $pPhpPresentation
     * @return \PhpOffice\PhpPresentation\Shape\AbstractDrawing[] All drawings in PhpPresentation
     * @throws \Exception
     */
    public function allDrawings(PhpPresentation $pPhpPresentation)
    {
        // Get an array of all drawings
        $aDrawings  = array();

        // Loop trough PhpPresentation
        $slideCount = $pPhpPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            // Loop trough images and add to array
            $iterator = $pPhpPresentation->getSlide($i)->getShapeCollection()->getIterator();
            while ($iterator->valid()) {
                if ($iterator->current() instanceof AbstractDrawing && !($iterator->current() instanceof Table)) {
                    $aDrawings[] = $iterator->current();
                } elseif ($iterator->current() instanceof Group) {
                    $iterator2 = $iterator->current()->getShapeCollection()->getIterator();
                    while ($iterator2->valid()) {
                        if ($iterator2->current() instanceof AbstractDrawing && !($iterator2->current() instanceof Table)) {
                            $aDrawings[] = $iterator2->current();
                        }
                        $iterator2->next();
                    }
                }

                $iterator->next();
            }
        }

        return $aDrawings;
    }
}
