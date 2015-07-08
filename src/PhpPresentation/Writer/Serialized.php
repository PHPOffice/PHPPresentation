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

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AbstractDrawing;

/**
 * \PhpOffice\PhpPresentation\Writer\Serialized
 */
class Serialized implements WriterInterface
{
    /**
     * Private PhpPresentation
     *
     * @var \PhpOffice\PhpPresentation\PhpPresentation
     */
    private $presentation;

    /**
     * Create a new \PhpOffice\PhpPresentation\Writer\Serialized
     *
     * @param \PhpOffice\PhpPresentation\PhpPresentation $pPhpPresentation
     */
    public function __construct(PhpPresentation $pPhpPresentation = null)
    {
        // Assign PhpPresentation
        $this->setPhpPresentation($pPhpPresentation);
    }

    /**
     * Save PhpPresentation to file
     *
     * @param  string    $pFilename
     * @throws \Exception
     */
    public function save($pFilename)
    {
        if (empty($pFilename)) {
            throw new \Exception("Filename is empty.");
        }
        if (!is_null($this->presentation)) {
            // Create new ZIP file and open it for writing
            $objZip = new \ZipArchive();

            // Try opening the ZIP file
            if ($objZip->open($pFilename, \ZipArchive::CREATE) !== true) {
                if ($objZip->open($pFilename, \ZipArchive::OVERWRITE) !== true) {
                    throw new \Exception("Could not open " . $pFilename . " for writing.");
                }
            }

            // Add media
            $slideCount = $this->presentation->getSlideCount();
            for ($i = 0; $i < $slideCount; ++$i) {
                for ($j = 0; $j < $this->presentation->getSlide($i)->getShapeCollection()->count(); ++$j) {
                    if ($this->presentation->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof AbstractDrawing) {
                        $imgTemp = $this->presentation->getSlide($i)->getShapeCollection()->offsetGet($j);
                        $objZip->addFromString('media/' . $imgTemp->getFilename(), file_get_contents($imgTemp->getPath()));
                    }
                }
            }

            // Add PhpPresentation.xml to the document, which represents a PHP serialized PhpPresentation object
            $objZip->addFromString('PhpPresentation.xml', $this->writeSerialized($this->presentation, $pFilename));

            // Close file
            if ($objZip->close() === false) {
                throw new \Exception("Could not close zip file $pFilename.");
            }
        } else {
            throw new \Exception("PhpPresentation object unassigned.");
        }
    }

    /**
     * Get PhpPresentation object
     *
     * @return PhpPresentation
     * @throws \Exception
     */
    public function getPhpPresentation()
    {
        if (!is_null($this->presentation)) {
            return $this->presentation;
        } else {
            throw new \Exception("No PhpPresentation assigned.");
        }
    }

    /**
     * Get PhpPresentation object
     *
     * @param  PhpPresentation                   $pPhpPresentation PhpPresentation object
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Writer\Serialized
     */
    public function setPhpPresentation(PhpPresentation $pPhpPresentation = null)
    {
        $this->presentation = $pPhpPresentation;

        return $this;
    }

    /**
     * Serialize PhpPresentation object to XML
     *
     * @param  PhpPresentation $pPhpPresentation
     * @param  string        $pFilename
     * @return string        XML Output
     * @throws \Exception
     */
    private function writeSerialized(PhpPresentation $pPhpPresentation = null, $pFilename = '')
    {
        // Clone $pPhpPresentation
        $pPhpPresentation = clone $pPhpPresentation;

        // Update media links
        $slideCount = $pPhpPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            for ($j = 0; $j < $pPhpPresentation->getSlide($i)->getShapeCollection()->count(); ++$j) {
                if ($pPhpPresentation->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof AbstractDrawing) {
                    $pPhpPresentation->getSlide($i)->getShapeCollection()->offsetGet($j)->setPath('zip://' . $pFilename . '#media/' . $pPhpPresentation->getSlide($i)->getShapeCollection()->offsetGet($j)->getFilename(), false);
                }
            }
        }

        // Create XML writer
        $objWriter = new XMLWriter();
        $objWriter->openMemory();
        $objWriter->setIndent(true);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // PhpPresentation
        $objWriter->startElement('PhpPresentation');
        $objWriter->writeAttribute('version', '##VERSION##');

        // Comment
        $objWriter->writeComment('This file has been generated using PhpPresentation v##VERSION## (http://github.com/PHPOffice/PhpPresentation). It contains a base64 encoded serialized version of the PhpPresentation internal object.');

        // Data
        $objWriter->startElement('data');
        $objWriter->writeCData(base64_encode(serialize($pPhpPresentation)));
        $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
