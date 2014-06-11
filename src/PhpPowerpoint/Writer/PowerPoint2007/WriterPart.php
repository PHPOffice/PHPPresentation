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

use PhpOffice\PhpPowerpoint\Writer\IWriter;

/**
 * PHPPowerPoint_Writer_PowerPoint2007_WriterPart
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
abstract class WriterPart
{
    /**
     * Parent IWriter object
     *
     * @var PHPPowerPoint_Writer_IWriter
     */
    private $_parentWriter;

    /**
     * Set parent IWriter object
     *
     * @param  PHPPowerPoint_Writer_IWriter $pWriter
     * @throws Exception
     */
    public function setParentWriter(IWriter $pWriter = null)
    {
        $this->_parentWriter = $pWriter;
    }

    /**
     * Get parent IWriter object
     *
     * @return PHPPowerPoint_Writer_IWriter
     * @throws Exception
     */
    public function getParentWriter()
    {
        if (!is_null($this->_parentWriter)) {
            return $this->_parentWriter;
        } else {
            throw new Exception("No parent PHPPowerPoint_Writer_IWriter assigned.");
        }
    }
}
