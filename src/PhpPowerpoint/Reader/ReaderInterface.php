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

namespace PhpOffice\PhpPowerpoint\Reader;

/**
 * Reader interface
 */
interface ReaderInterface
{
    /**
     * Can the current \PhpOffice\PHPPowerPoint\Reader\ReaderInterface read the file?
     *
     * @param  string  $pFilename
     * @return boolean
     */
    public function canRead($pFilename);

    /**
     * Loads PHPPowerPoint from file
     *
     * @param  string    $pFilename
     * @throws \Exception
     */
    public function load($pFilename);
}
