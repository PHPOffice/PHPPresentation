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
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;
use PhpOffice\PhpPowerpoint\Writer\WriterInterface;
use PhpOffice\PhpPowerpoint\Writer\ODPresentation;

/**
 * \PhpOffice\PhpPowerpoint\Writer\ODPresentation\AbstractPart
 */
abstract class AbstractPart
{
//     abstract public function writePart(PhpPowerpoint $pPHPPowerPoint);
    
    /**
     * Parent WriterInterface object
     *
     * @var \PhpOffice\PhpPowerpoint\Writer\WriterInterface
     */
    private $parentWriter;

    /**
     * Set parent WriterInterface object
     *
     * @param  \PhpOffice\PhpPowerpoint\Writer\WriterInterface $pWriter
     * @throws \Exception
     */
    public function setParentWriter(WriterInterface $pWriter = null)
    {
        $this->parentWriter = $pWriter;
    }

    /**
     * Get parent WriterInterface object
     *
     * @return \PhpOffice\PhpPowerpoint\Writer\WriterInterface
     * @throws \Exception
     */
    public function getParentWriter()
    {
        if (!is_null($this->parentWriter)) {
            return $this->parentWriter;
        } else {
            throw new \Exception("No parent \PhpOffice\PhpPowerpoint\Writer\WriterInterface assigned.");
        }
    }

    /**
     * Get XML Writer
     */
    protected function getXMLWriter()
    {
        $parentWriter = $this->getParentWriter();
        if (!$parentWriter instanceof ODPresentation) {
            throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPowerpoint\Writer\ODPresentation');
        }
        if ($parentWriter->hasDiskCaching()) {
            return new XMLWriter(XMLWriter::STORAGE_DISK, $parentWriter->getDiskCachingDirectory());
        } else {
            return new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }
    }
}
