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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Writer\WriterInterface;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007\AbstractPart
 */
abstract class AbstractPart
{
    /**
     * Parent WriterInterface object
     *
     * @var \PhpOffice\PhpPresentation\Writer\WriterInterface
     */
    private $parentWriter;

    /**
     * Set parent WriterInterface object
     *
     * @param  \PhpOffice\PhpPresentation\Writer\WriterInterface $pWriter
     * @throws \Exception
     */
    public function setParentWriter(WriterInterface $pWriter = null)
    {
        $this->parentWriter = $pWriter;
    }

    /**
     * Get parent WriterInterface object
     *
     * @return \PhpOffice\PhpPresentation\Writer\WriterInterface
     * @throws \Exception
     */
    public function getParentWriter()
    {
        if (!is_null($this->parentWriter)) {
            return $this->parentWriter;
        } else {
            throw new \Exception("No parent \PhpOffice\PhpPresentation\Writer\WriterInterface assigned.");
        }
    }

    /**
     * Get XML writer
     */
    protected function getXMLWriter()
    {
        $parentWriter = $this->getParentWriter();
        if (!$parentWriter instanceof PowerPoint2007) {
            throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\PowerPoint2007');
        }
        if ($parentWriter->hasDiskCaching()) {
            return new XMLWriter(XMLWriter::STORAGE_DISK, $parentWriter->getDiskCachingDirectory());
        } else {
            return new XMLWriter(XMLWriter::STORAGE_MEMORY);
        }
    }
}
