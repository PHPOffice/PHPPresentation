<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Writer_PowerPoint2007_Drawing
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_Drawing extends PHPPowerPoint_Writer_PowerPoint2007_WriterPart
{
	/**
	 * Get an array of all drawings
	 *
	 * @param 	PHPPowerPoint							$pPHPPowerPoint
	 * @return 	PHPPowerPoint_Slide_Drawing[]		All drawings in PHPPowerPoint
	 * @throws 	Exception
	 */
	public function allDrawings(PHPPowerPoint $pPHPPowerPoint = null)
	{
		// Get an array of all drawings
		$aDrawings	= array();

		// Loop trough PHPPowerPoint
		$slideCount = $pPHPPowerPoint->getSlideCount();
		for ($i = 0; $i < $slideCount; ++$i) {
			// Loop trough images and add to array
			$iterator = $pPHPPowerPoint->getSlide($i)->getShapeCollection()->getIterator();
			while ($iterator->valid()) {
				if ($iterator->current() instanceof PHPPowerPoint_Shape_BaseDrawing && !($iterator->current() instanceof PHPPowerPoint_Shape_Table)) {
					$aDrawings[] = $iterator->current();
				}

  				$iterator->next();
			}
		}

		return $aDrawings;
	}
}
