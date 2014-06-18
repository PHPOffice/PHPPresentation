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

use PhpOffice\PhpPowerpoint\Writer\IWriter;
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;

/**
 * PHPPowerPoint_Writer_ODPresentation_WriterPart
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
abstract class WriterPart
{
    /**
     * Parent IWriter object
     *
     * @var PHPPowerPoint_Writer_IWriter
     */
    private $parentWriter;

    /**
     * Set parent IWriter object
     *
     * @param  PHPPowerPoint_Writer_IWriter $pWriter
     * @throws \Exception
     */
    public function setParentWriter(IWriter $pWriter = null)
    {
        $this->parentWriter = $pWriter;
    }

    /**
     * Get parent IWriter object
     *
     * @return PHPPowerPoint_Writer_IWriter
     * @throws \Exception
     */
    public function getParentWriter()
    {
        if (!is_null($this->parentWriter)) {
            return $this->parentWriter;
        } else {
            throw new \Exception("No parent PHPPowerPoint_Writer_IWriter assigned.");
        }
    }
    
    protected function getXMLWriter()
    {
    	if ($this->getParentWriter()->hasDiskCaching()) {
    		return new XMLWriter(XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
    	} else {
    		return new XMLWriter(XMLWriter::STORAGE_MEMORY);
    	}
    }
}
