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

use PhpOffice\PhpPowerpoint\Writer\ODPresentation\WriterPart;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\BaseDrawing;
use PhpOffice\PhpPowerpoint\Shape\Table;

/**
 * PHPPowerPoint_Writer_ODPresentation_Drawing
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Drawing extends WriterPart
{
    /**
     * Get an array of all drawings
     *
     * @param  PHPPowerPoint                 $pPHPPowerPoint
     * @return PHPPowerPoint_Slide_Drawing[] All drawings in PHPPowerPoint
     * @throws Exception
     */
    public function allDrawings(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Get an array of all drawings
        $aDrawings  = array();

        // Loop trough PHPPowerPoint
        $slideCount = $pPHPPowerPoint->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            // Loop trough images and add to array
            $iterator = $pPHPPowerPoint->getSlide($i)->getShapeCollection()->getIterator();
            while ($iterator->valid()) {
                if ($iterator->current() instanceof BaseDrawing && !($iterator->current() instanceof Table)) {
                    $aDrawings[] = $iterator->current();
                }

                $iterator->next();
            }
        }

        return $aDrawings;
    }
}
