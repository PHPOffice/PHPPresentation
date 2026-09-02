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
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Exception\DirectoryNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Reader\Keynote as KeynoteReader;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide;

/**
 * Keynote writer : the text and the images of a presentation, as an Apple Keynote package.
 *
 * The package is written in the shape Keynote '09 reads : a Zip archive of an `index.apxl` XML
 * document and of the images the slides use. The IWA shape of Keynote '13 and later is not
 * written -- it would mean writing Apple's Protobuf schemas, and the request this answers asks
 * only for text and images -- but it is the shape the matching reader prefers when it finds it.
 */
class Keynote extends AbstractWriter implements WriterInterface
{
    /**
     * The folder of the package the images are written to.
     *
     * @var string
     */
    protected const PATH_DATA = 'Data/';

    /**
     * Create a new \PhpOffice\PhpPresentation\Writer\Keynote.
     */
    public function __construct(?PhpPresentation $pPhpPresentation = null)
    {
        $this->setPhpPresentation($pPhpPresentation ?? new PhpPresentation());
        $this->oDrawingHashTable = new HashTable();
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

        $this->getDrawingHashTable()->addFromSource($this->allDrawings());

        $oZip = $this->getZipAdapter();
        $oZip->open($pFilename);
        $oZip->addFromString('index.apxl', $this->writeApxl());
        foreach ($this->getPhpPresentation()->getAllSlides() as $oSlide) {
            foreach ($oSlide->getShapeCollection() as $oShape) {
                if ($oShape instanceof AbstractDrawingAdapter) {
                    $oZip->addFromString(self::PATH_DATA . $oShape->getIndexedFilename(), $oShape->getContents());
                }
            }
        }
        $oZip->close();
    }

    /**
     * The `index.apxl` document of the package.
     */
    protected function writeApxl(): string
    {
        $oPresentation = $this->getPhpPresentation();
        $oLayout = $oPresentation->getLayout();

        $objWriter = new XMLWriter();
        $objWriter->openMemory();
        $objWriter->setIndent(true);
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        $objWriter->startElement('key:presentation');
        $objWriter->writeAttribute('xmlns:key', KeynoteReader::NS_KEY);
        $objWriter->writeAttribute('xmlns:sf', KeynoteReader::NS_SF);
        $objWriter->writeAttribute('xmlns:sfa', KeynoteReader::NS_SFA);
        $objWriter->writeAttribute('key:version', '92008102400');

        $objWriter->startElement('key:size');
        $objWriter->writeAttribute('sfa:w', (string) round($oLayout->getCX(DocumentLayout::UNIT_PIXEL)));
        $objWriter->writeAttribute('sfa:h', (string) round($oLayout->getCY(DocumentLayout::UNIT_PIXEL)));
        $objWriter->endElement();

        $objWriter->startElement('key:slide-list');
        foreach ($oPresentation->getAllSlides() as $oSlide) {
            $this->writeSlide($objWriter, $oSlide);
        }
        $objWriter->endElement();

        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * One slide of the presentation, with its drawables and its speaker note.
     */
    protected function writeSlide(XMLWriter $objWriter, Slide $oSlide): void
    {
        $objWriter->startElement('key:slide');

        $objWriter->startElement('key:drawables');
        foreach ($oSlide->getShapeCollection() as $oShape) {
            if ($oShape instanceof RichText) {
                $this->writeText($objWriter, $oShape);
            }
            if ($oShape instanceof AbstractDrawingAdapter) {
                $this->writeImage($objWriter, $oShape);
            }
        }
        $objWriter->endElement();

        $notes = [];
        foreach ($oSlide->getNote()->getShapeCollection() as $oShape) {
            if ($oShape instanceof RichText) {
                $notes = array_merge($notes, $this->getParagraphs($oShape));
            }
        }
        if ([] !== $notes) {
            $objWriter->startElement('key:notes');
            $this->writeTextStorage($objWriter, $notes);
            $objWriter->endElement();
        }

        $objWriter->endElement();
    }

    /**
     * A text shape, as the placeholder Keynote stores a body of text in.
     */
    protected function writeText(XMLWriter $objWriter, RichText $oShape): void
    {
        $objWriter->startElement('key:body-placeholder');
        $this->writeGeometry($objWriter, $oShape->getOffsetX(), $oShape->getOffsetY(), $oShape->getWidth(), $oShape->getHeight());
        $this->writeTextStorage($objWriter, $this->getParagraphs($oShape));
        $objWriter->endElement();
    }

    /**
     * An image of a slide, pointing at the file of the package which holds it.
     */
    protected function writeImage(XMLWriter $objWriter, AbstractDrawingAdapter $oShape): void
    {
        $filename = $oShape->getIndexedFilename();

        $objWriter->startElement('key:image');
        $this->writeGeometry($objWriter, $oShape->getOffsetX(), $oShape->getOffsetY(), $oShape->getWidth(), $oShape->getHeight());
        $objWriter->startElement('sf:content');
        $objWriter->startElement('sf:image-media');
        $objWriter->startElement('sf:filtered-image');
        $objWriter->startElement('sf:unfiltered');
        $objWriter->startElement('sf:size');
        $objWriter->writeAttribute('sfa:w', (string) $oShape->getWidth());
        $objWriter->writeAttribute('sfa:h', (string) $oShape->getHeight());
        $objWriter->endElement();
        $objWriter->startElement('sf:data');
        $objWriter->writeAttribute('sf:path', self::PATH_DATA . $filename);
        $objWriter->writeAttribute('sf:displayname', $filename);
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * The position and the size of a drawable.
     */
    protected function writeGeometry(XMLWriter $objWriter, int $offsetX, int $offsetY, int $width, int $height): void
    {
        $objWriter->startElement('sf:geometry');
        $objWriter->startElement('sf:position');
        $objWriter->writeAttribute('sfa:x', (string) $offsetX);
        $objWriter->writeAttribute('sfa:y', (string) $offsetY);
        $objWriter->endElement();
        $objWriter->startElement('sf:size');
        $objWriter->writeAttribute('sfa:w', (string) $width);
        $objWriter->writeAttribute('sfa:h', (string) $height);
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * Paragraphs, as the text storage Keynote wraps them in.
     *
     * @param array<int, string> $paragraphs
     */
    protected function writeTextStorage(XMLWriter $objWriter, array $paragraphs): void
    {
        $objWriter->startElement('sf:text');
        $objWriter->startElement('sf:text-storage');
        $objWriter->startElement('sf:text-body');
        foreach ($paragraphs as $paragraph) {
            $objWriter->startElement('sf:p');
            $objWriter->writeElement('sf:span', $paragraph);
            $objWriter->endElement();
        }
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
    }

    /**
     * The paragraphs of a text shape, as plain text.
     *
     * @return array<int, string>
     */
    protected function getParagraphs(RichText $oShape): array
    {
        $paragraphs = [];
        foreach ($oShape->getParagraphs() as $oParagraph) {
            $paragraphs[] = $oParagraph->getPlainText();
        }

        return '' === trim(implode('', $paragraphs)) ? [] : $paragraphs;
    }
}
