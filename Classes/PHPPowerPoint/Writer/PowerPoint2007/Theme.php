<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2006 - 2009 PHPPowerPoint
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
 * PHPPowerPoint_Writer_PowerPoint2007_Theme
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2006 - 2009 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_Theme extends PHPPowerPoint_Writer_PowerPoint2007_WriterPart
{
	/**
	 * Write theme to XML format
	 *
	 * @param 	PHPPowerPoint	$pPHPPowerPoint
	 * @param	int				$masterId
	 * @return 	string 			XML Output
	 * @throws 	Exception
	 */
	public function writeTheme(PHPPowerPoint $pPHPPowerPoint = null, $masterId = 1)
	{
		// Write theme from layout pack
		$layoutPack		= $this->getParentWriter()->getLayoutPack();
		foreach ($layoutPack->getThemes() as $theme) {
			if ($theme['masterid'] == $masterId) {
				return $theme['body'];
			}
		}
		throw new Exception('No theme has been found!');
	}
}
