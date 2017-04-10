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

use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;

/**
 * \PhpOffice\PhpPresentation\Writer\Serialized
 */
class Serialized extends AbstractWriter implements WriterInterface
{
    /**
     * Create a new \PhpOffice\PhpPresentation\Writer\Serialized
     *
     * @param \PhpOffice\PhpPresentation\PhpPresentation $pPhpPresentation
     */
    public function __construct(PhpPresentation $pPhpPresentation = null)
    {
        // Set PhpPresentation
        $this->setPhpPresentation($pPhpPresentation);

        // Set ZIP Adapter
        $this->setZipAdapter(new ZipArchiveAdapter());
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
        $oPresentation = $this->getPhpPresentation();

        // Create new ZIP file and open it for writing
        $objZip = $this->getZipAdapter();

        // Try opening the ZIP file
        $objZip->open($pFilename);

        // Add media
        $slideCount = $oPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            for ($j = 0; $j < $oPresentation->getSlide($i)->getShapeCollection()->count(); ++$j) {
                if ($oPresentation->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof AbstractDrawingAdapter) {
                    $imgTemp = $oPresentation->getSlide($i)->getShapeCollection()->offsetGet($j);
                    $objZip->addFromString('media/' . $imgTemp->getIndexedFilename(), file_get_contents($imgTemp->getPath()));
                }
            }
        }

        // Add PhpPresentation.xml to the document, which represents a PHP serialized PhpPresentation object
        $objZip->addFromString('PhpPresentation.xml', $this->writeSerialized($oPresentation, $pFilename));

        // Close file
        $objZip->close();
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
                if ($pPhpPresentation->getSlide($i)->getShapeCollection()->offsetGet($j) instanceof AbstractDrawingAdapter) {
                    $pPhpPresentation->getSlide($i)->getShapeCollection()->offsetGet($j)->setPath('zip://' . $pFilename . '#media/' . $pPhpPresentation->getSlide($i)->getShapeCollection()->offsetGet($j)->getIndexedFilename(), false);
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
