<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPowerpoint\Writer\ODPresentation;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\AbstractDrawing;
use PhpOffice\PhpPowerpoint\Shape\Group;
use PhpOffice\PhpPowerpoint\Shape\Table;

/**
 * \PhpOffice\PhpPowerpoint\Writer\ODPresentation\Drawing
 */
class Drawing extends AbstractPart
{
    /**
     * Get an array of all drawings
     *
     * @param  PhpPowerPoint                 $pPHPPowerPoint
     * @return \PhpOffice\PhpPowerpoint\Shape\AbstractDrawing[] All drawings in PHPPowerPoint
     * @throws \Exception
     */
    public function allDrawings(PhpPowerPoint $pPHPPowerPoint)
    {
        // Get an array of all drawings
        $aDrawings  = array();

        // Loop trough PHPPowerPoint
        $slideCount = $pPHPPowerPoint->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            // Loop trough images and add to array
            $iterator = $pPHPPowerPoint->getSlide($i)->getShapeCollection()->getIterator();
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
