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

namespace PhpOffice\PhpPowerpoint\Writer;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\AbstractDrawing;
use PhpOffice\PhpPowerpoint\Shared\XMLWriter;

/**
 * \PhpOffice\PhpPowerpoint\Writer\Serialized
 */
class Serialized implements WriterInterface
{
    /**
     * Private PHPPowerPoint
     *
     * @var \PhpOffice\PhpPowerpoint\PhpPowerpoint
     */
    private $presentation;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Writer\Serialized
     *
     * @param \PhpOffice\PhpPowerpoint\PhpPowerpoint $pPHPPowerPoint
     */
    public function __construct (PhpPowerpoint $pPHPPowerPoint = null)
    {
        // Assign PHPPowerPoint
        $this->setPHPPowerPoint($pPHPPowerPoint);
    }

    /**
     * Save PHPPowerPoint to file
     *
     * @param  string    $pFilename
     * @throws \Exception
     */
    public function save ($pFilename)
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

            // Add PHPPowerPoint.xml to the document, which represents a PHP serialized PHPPowerPoint object
            $objZip->addFromString('PHPPowerPoint.xml', $this->writeSerialized($this->presentation, $pFilename));

            // Close file
            if ($objZip->close() === false) {
                throw new \Exception("Could not close zip file $pFilename.");
            }
        } else {
            throw new \Exception("PHPPowerPoint object unassigned.");
        }
    }

    /**
     * Get PHPPowerPoint object
     *
     * @return PHPPowerPoint
     * @throws \Exception
     */
    public function getPHPPowerPoint ()
    {
        if (!is_null($this->presentation)) {
            return $this->presentation;
        } else {
            throw new \Exception("No PHPPowerPoint assigned.");
        }
    }

    /**
     * Get PHPPowerPoint object
     *
     * @param  PHPPowerPoint                   $pPHPPowerPoint PHPPowerPoint object
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Writer\Serialized
     */
    public function setPHPPowerPoint (PhpPowerpoint $pPHPPowerPoint = null)
    {
        $this->presentation = $pPHPPowerPoint;

        return $this;
    }

    /**
     * Serialize PHPPowerPoint object to XML
     *
     * @param  PHPPowerPoint $pPHPPowerPoint
     * @param  string        $pFilename
     * @return string        XML Output
     * @throws \Exception
     */
    private function writeSerialized (PhpPowerpoint $pPHPPowerPoint = null, $pFilename = '')
    {
        // Clone $pPHPPowerPoint
        $pPHPPowerPoint = clone $pPHPPowerPoint;

        // Update media links
        $slideCount = $pPHPPowerPoint->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            for ($j = 0; $j < $pPHPPowerPoint->getSlide($i)->getShapeCollection()->count(); ++$j) {
                if ($pPHPPowerPoint->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof AbstractDrawing) {
                    $pPHPPowerPoint->getSlide($i)->getShapeCollection()->offsetGet($j)->setPath('zip://' . $pFilename . '#media/' . $pPHPPowerPoint->getSlide($i)->getShapeCollection()->offsetGet($j)->getFilename(), false);
                }
            }
        }

        // Create XML writer
        $objWriter = new XMLWriter();
        $objWriter->openMemory();
        $objWriter->setIndent(true);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // PHPPowerPoint
        $objWriter->startElement('PHPPowerPoint');
        $objWriter->writeAttribute('version', '##VERSION##');

        // Comment
        $objWriter->writeComment('This file has been generated using PHPPowerPoint v##VERSION## (http://github.com/PHPOffice/PHPPowerPoint). It contains a base64 encoded serialized version of the PHPPowerPoint internal object.');

        // Data
        $objWriter->startElement('data');
        $objWriter->writeCData(base64_encode(serialize($pPHPPowerPoint)));
        $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
