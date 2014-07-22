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

namespace PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;

use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007;

/**
 * \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\Theme
 */
class Theme extends AbstractPart
{
    /**
     * Write theme to XML format
     *
     * @param  int $masterId
     * @throws \Exception
     * @return string XML Output
     */
    public function writeTheme($masterId = 1)
    {
        // Write theme from layout pack
        $parentWriter = $this->getParentWriter();
        if ($parentWriter instanceof PowerPoint2007) {
            $layoutPack     = $parentWriter->getLayoutPack();
            foreach ($layoutPack->getThemes() as $theme) {
                if ($theme['masterid'] == $masterId) {
                    return $theme['body'];
                }
            }
            throw new \Exception('No theme has been found!');
        }
    }
}
