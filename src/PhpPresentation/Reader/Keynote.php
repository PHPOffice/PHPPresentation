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

namespace PhpOffice\PhpPresentation\Reader;

use DOMElement;
use PhpOffice\Common\XMLReader;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide;
use ZipArchive;

/**
 * Reader for the Apple Keynote (.key) presentation format.
 *
 * This reader targets the legacy iWork '09 single-file structure, where the
 * presentation is described by an `index.apxl` XML document stored inside the
 * Keynote package (a Zip archive). The modern IWA binary format (Keynote '13 and
 * later) is intentionally not handled and is safely rejected by canRead().
 */
class Keynote implements ReaderInterface
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
     * Output object.
     *
     * @var PhpPresentation
     */
    protected $oPhpPresentation;

    /**
     * Should the images be loaded?
     *
     * @var bool
     */
    protected $loadImages = true;

    /**
     * Can the current reader read the file?
     */
    public function canRead(string $pFilename): bool
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support being read by this reader?
     */
    public function fileSupportsUnserializePhpPresentation(string $pFilename = ''): bool
    {
        if (!file_exists($pFilename)) {
            throw new FileNotFoundException($pFilename);
        }

        $oZip = new ZipArchive();
        if (true === $oZip->open($pFilename)) {
            $isApxl = is_array($oZip->statName('index.apxl'));
            $oZip->close();

            return $isApxl;
        }

        return false;
    }

    /**
     * Loads a Keynote file.
     */
    public function load(string $pFilename, int $flags = 0): PhpPresentation
    {
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new InvalidFileFormatException($pFilename, self::class);
        }

        $this->loadImages = !((bool) ($flags & self::SKIP_IMAGES));

        return $this->loadFile($pFilename);
    }

    /**
     * Loads the file content into a PhpPresentation object.
     */
    protected function loadFile(string $pFilename): PhpPresentation
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();

        $oXMLReader = new XMLReader();
        if (false !== $oXMLReader->getDomFromZip($pFilename, 'index.apxl')) {
            $oXMLReader->registerNamespace('key', self::NS_KEY);
            $oXMLReader->registerNamespace('sf', self::NS_SF);
            $oXMLReader->registerNamespace('sfa', self::NS_SFA);
            $this->loadSlides($oXMLReader);
        }

        return $this->oPhpPresentation;
    }

    /**
     * Reads every slide of the presentation.
     */
    protected function loadSlides(XMLReader $oXMLReader): void
    {
        foreach ($oXMLReader->getElements('/key:presentation/key:slide-list/key:slide') as $oNodeSlide) {
            if (!$oNodeSlide instanceof DOMElement) {
                continue;
            }
            $oSlide = $this->oPhpPresentation->createSlide();
            $this->loadShapes($oXMLReader, $oNodeSlide, $oSlide);
            $this->loadNote($oXMLReader, $oNodeSlide, $oSlide);
        }
    }

    /**
     * Reads every text placeholder of a slide.
     */
    protected function loadShapes(XMLReader $oXMLReader, DOMElement $oNodeSlide, Slide $oSlide): void
    {
        foreach ($oXMLReader->getElements('./key:drawables/key:body-placeholder', $oNodeSlide) as $oNodeShape) {
            if (!$oNodeShape instanceof DOMElement) {
                continue;
            }
            $oRichText = $oSlide->createRichTextShape();
            $this->loadGeometry($oXMLReader, $oNodeShape, $oRichText);
            $this->loadText($oXMLReader, $oNodeShape, $oRichText);
        }
    }

    /**
     * Reads the position and size of a shape.
     */
    protected function loadGeometry(XMLReader $oXMLReader, DOMElement $oNodeShape, RichText $oRichText): void
    {
        $oNodePosition = $oXMLReader->getElement('./sf:geometry/sf:position', $oNodeShape);
        if ($oNodePosition instanceof DOMElement) {
            $oRichText->setOffsetX((int) $oNodePosition->getAttributeNS(self::NS_SFA, 'x'));
            $oRichText->setOffsetY((int) $oNodePosition->getAttributeNS(self::NS_SFA, 'y'));
        }

        $oNodeSize = $oXMLReader->getElement('./sf:geometry/sf:size', $oNodeShape);
        if ($oNodeSize instanceof DOMElement) {
            $oRichText->setWidth((int) $oNodeSize->getAttributeNS(self::NS_SFA, 'w'));
            $oRichText->setHeight((int) $oNodeSize->getAttributeNS(self::NS_SFA, 'h'));
        }
    }

    /**
     * Reads the paragraphs and runs of a text container.
     */
    protected function loadText(XMLReader $oXMLReader, DOMElement $oNodeContainer, RichText $oRichText): void
    {
        $oNodeBody = $oXMLReader->getElement('./sf:text/sf:text-storage/sf:text-body', $oNodeContainer);
        if (!$oNodeBody instanceof DOMElement) {
            return;
        }

        $isFirstParagraph = true;
        foreach ($oXMLReader->getElements('./sf:p', $oNodeBody) as $oNodeParagraph) {
            if (!$oNodeParagraph instanceof DOMElement) {
                continue;
            }
            $oParagraph = $isFirstParagraph
                ? $oRichText->getActiveParagraph()
                : $oRichText->createParagraph();
            $isFirstParagraph = false;

            foreach ($oXMLReader->getElements('./sf:span', $oNodeParagraph) as $oNodeSpan) {
                if (!$oNodeSpan instanceof DOMElement) {
                    continue;
                }
                $oParagraph->createTextRun((string) $oNodeSpan->nodeValue);
            }
        }
    }

    /**
     * Reads the speaker note of a slide.
     */
    protected function loadNote(XMLReader $oXMLReader, DOMElement $oNodeSlide, Slide $oSlide): void
    {
        $oNodeNote = $oXMLReader->getElement('./key:notes', $oNodeSlide);
        if (!$oNodeNote instanceof DOMElement) {
            return;
        }

        $oRichText = $oSlide->getNote()->createRichTextShape();
        $this->loadText($oXMLReader, $oNodeNote, $oRichText);
    }
}
