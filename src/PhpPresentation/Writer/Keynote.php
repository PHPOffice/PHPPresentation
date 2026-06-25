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
use PhpOffice\PhpPresentation\Exception\FileCopyException;
use PhpOffice\PhpPresentation\Exception\FileRemoveException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Slide;

/**
 * Writer for the Apple Keynote (.key) presentation format.
 *
 * The writer produces a Keynote package (a Zip archive) holding an `index.apxl`
 * XML document that follows the legacy iWork '09 structure. The output is meant to
 * be round-trippable with the matching reader; opening it directly in a current
 * Keynote.app is not guaranteed because of the format generation gap.
 *
 * Only text placeholders (their paragraphs and text runs), shape geometry and
 * speaker notes are written. Other content (images, tables, charts and line
 * breaks) is not represented in this basic format and is silently skipped on save.
 */
class Keynote extends AbstractWriter implements WriterInterface
{
    /**
     * Keynote namespace.
     *
     * @var string
     */
    public const NS_KEY = 'http://developer.apple.com/namespaces/keynote2';

    /**
     * iWork shared namespace.
     *
     * @var string
     */
    public const NS_SF = 'http://developer.apple.com/namespaces/sf';

    /**
     * iWork shared abstract namespace.
     *
     * @var string
     */
    public const NS_SFA = 'http://developer.apple.com/namespaces/sfa';

    /**
     * Create a new Keynote writer.
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

        // If $pFilename is php://output or php://stdout, make it a temporary file...
        $originalFilename = $pFilename;
        if ('php://output' == strtolower($pFilename) || 'php://stdout' == strtolower($pFilename)) {
            $pFilename = @tempnam(sys_get_temp_dir(), 'phppttmp');
            if (false === $pFilename) {
                $pFilename = $originalFilename;
            }
        }

        $oZip = $this->getZipAdapter();
        $oZip->open($pFilename);
        $oZip->addFromString('index.apxl', $this->writeContent());
        $oZip->close();

        // If a temporary file was used, copy it to the correct file stream
        if ($originalFilename != $pFilename) {
            if (false === copy($pFilename, $originalFilename)) {
                throw new FileCopyException($pFilename, $originalFilename);
            }
            if (false === @unlink($pFilename)) {
                throw new FileRemoveException($pFilename);
            }
        }
    }

    /**
     * Builds the index.apxl document.
     */
    protected function writeContent(): string
    {
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        $objWriter->startDocument('1.0', 'UTF-8');

        $objWriter->startElement('key:presentation');
        $objWriter->writeAttribute('xmlns:key', self::NS_KEY);
        $objWriter->writeAttribute('xmlns:sf', self::NS_SF);
        $objWriter->writeAttribute('xmlns:sfa', self::NS_SFA);
        $objWriter->writeAttribute('key:version', '92008102400');

        $objWriter->startElement('key:size');
        $objWriter->writeAttribute('sfa:w', '960');
        $objWriter->writeAttribute('sfa:h', '720');
        $objWriter->endElement();

        $objWriter->startElement('key:theme-list');
        $objWriter->startElement('key:theme');
        $objWriter->writeAttribute('key:name', 'PHPPresentation');
        $objWriter->endElement();
        $objWriter->endElement();

        $objWriter->startElement('key:slide-list');
        foreach ($this->getPhpPresentation()->getAllSlides() as $oSlide) {
            $this->writeSlide($objWriter, $oSlide);
        }
        $objWriter->endElement();

        $objWriter->endElement();

        return $objWriter->getData();
    }

    /**
     * Writes a single slide.
     */
    protected function writeSlide(XMLWriter $objWriter, Slide $oSlide): void
    {
        $objWriter->startElement('key:slide');

        $objWriter->startElement('key:drawables');
        foreach ($oSlide->getShapeCollection() as $oShape) {
            if ($oShape instanceof RichText) {
                $this->writeShape($objWriter, $oShape);
            }
        }
        $objWriter->endElement();

        foreach ($oSlide->getNote()->getShapeCollection() as $oNoteShape) {
            if ($oNoteShape instanceof RichText) {
                $objWriter->startElement('key:notes');
                $this->writeText($objWriter, $oNoteShape);
                $objWriter->endElement();

                break;
            }
        }

        $objWriter->endElement();
    }

    /**
     * Writes a text placeholder shape, including its geometry.
     */
    protected function writeShape(XMLWriter $objWriter, RichText $oShape): void
    {
        $objWriter->startElement('key:body-placeholder');

        $objWriter->startElement('sf:geometry');
        $objWriter->startElement('sf:position');
        $objWriter->writeAttribute('sfa:x', (string) $oShape->getOffsetX());
        $objWriter->writeAttribute('sfa:y', (string) $oShape->getOffsetY());
        $objWriter->endElement();
        $objWriter->startElement('sf:size');
        $objWriter->writeAttribute('sfa:w', (string) $oShape->getWidth());
        $objWriter->writeAttribute('sfa:h', (string) $oShape->getHeight());
        $objWriter->endElement();
        $objWriter->endElement();

        $this->writeText($objWriter, $oShape);

        $objWriter->endElement();
    }

    /**
     * Writes the paragraphs and runs of a text container.
     */
    protected function writeText(XMLWriter $objWriter, RichText $oShape): void
    {
        $objWriter->startElement('sf:text');
        $objWriter->startElement('sf:text-storage');
        $objWriter->startElement('sf:text-body');
        foreach ($oShape->getParagraphs() as $oParagraph) {
            $objWriter->startElement('sf:p');
            foreach ($oParagraph->getRichTextElements() as $oElement) {
                // Only plain text runs are round-trippable; line breaks and other
                // element types are not represented in this basic format.
                if (!$oElement instanceof Run) {
                    continue;
                }
                $objWriter->startElement('sf:span');
                $objWriter->text((string) $oElement->getText());
                $objWriter->endElement();
            }
            $objWriter->endElement();
        }
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();
    }
}
