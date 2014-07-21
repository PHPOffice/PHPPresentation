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

/**
 * \PhpOffice\PhpPowerpoint\Writer\ODPresentation\Mimetype
 */
class Mimetype extends AbstractPart
{
    /**
     * Write Mimetype to Text format
     *
     * @return string        Text Output
     * @throws \Exception
     */
    public function writePart()
    {
        return 'application/vnd.oasis.opendocument.presentation';
    }
}
