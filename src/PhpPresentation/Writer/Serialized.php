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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Exception\DirectoryNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Drawing\File;

/**
 * \PhpOffice\PhpPresentation\Writer\Serialized.
 */
class Serialized extends AbstractWriter implements WriterInterface
{
    /**
     * Create a new \PhpOffice\PhpPresentation\Writer\Serialized.
     */
    public function __construct(?PhpPresentation $pPhpPresentation = null)
    {
        // Set PhpPresentation
        $this->setPhpPresentation($pPhpPresentation ?? new PhpPresentation());

        // Set ZIP Adapter
        $this->setZipAdapter(new ZipArchiveAdapter());
    }

    /**
     * Save PhpPresentation to file.
     */
    public function save(string $pFilename): void
    {
        if (empty($pFilename)) {
            throw new InvalidParameterException('pFilename', '');
        }
        if (!is_dir(dirname($pFilename))) {
            throw new DirectoryNotFoundException(dirname($pFilename));
        }
        $oPresentation = $this->getPhpPresentation();

        // Create new ZIP file and open it for writing
        $objZip = $this->getZipAdapter();

        // Try opening the ZIP file
        $objZip->open($pFilename);

        // Add media
        $slideCount = $oPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            foreach ($oPresentation->getSlide($i)->getShapeCollection() as $shape) {
                if ($shape instanceof AbstractDrawingAdapter) {
                    $objZip->addFromString(
                        'media/' . $shape->getImageIndex() . '/' . pathinfo($shape->getPath(), PATHINFO_BASENAME),
                        file_get_contents($shape->getPath())
                    );
                }
            }
        }

        // Add PhpPresentation.xml to the document, which represents a PHP serialized PhpPresentation object
        $objZip->addFromString('PhpPresentation.xml', $this->writeSerialized($oPresentation, $pFilename));

        // Close file
        $objZip->close();
    }

    /**
     * Serialize PhpPresentation object to XML.
     *
     * @param string $pFilename
     *
     * @return string XML Output
     */
    protected function writeSerialized(?PhpPresentation $pPhpPresentation = null, $pFilename = '')
    {
        // Clone $pPhpPresentation
        $pPhpPresentation = clone $pPhpPresentation;

        // Update media links
        $slideCount = $pPhpPresentation->getSlideCount();
        for ($i = 0; $i < $slideCount; ++$i) {
            foreach ($pPhpPresentation->getSlide($i)->getShapeCollection() as $shape) {
                if ($shape instanceof AbstractDrawingAdapter) {
                    $imgPath = 'zip://' . $pFilename . '#media/' . $shape->getImageIndex() . '/' . pathinfo($shape->getPath(), PATHINFO_BASENAME);
                    if ($shape instanceof File) {
                        $shape->setPath($imgPath, false);
                    } else {
                        $shape->setPath($imgPath);
                    }
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
