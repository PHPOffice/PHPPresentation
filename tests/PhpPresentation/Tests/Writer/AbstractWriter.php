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
 * @copyright   2009-2017 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests\Writer;

use PhpOffice\PhpPresentation\Writer;

/**
 * Mock class for AbstractWriter
 *
 */
class AbstractWriter extends Writer\AbstractWriter
{
    /**
     * public wrapper for protected method
     *
     * @return \PhpOffice\PhpPresentation\Shape\AbstractDrawing[] All drawings in PhpPresentation
     * @throws \Exception
     */
    public function allDrawings()
    {
        return parent::allDrawings();
    }
}
