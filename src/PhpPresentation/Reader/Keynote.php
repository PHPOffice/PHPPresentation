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
use PhpOffice\PhpPresentation\Reader\Keynote\IwaStream;
use PhpOffice\PhpPresentation\Reader\Keynote\Protobuf;
use PhpOffice\PhpPresentation\Reader\Keynote\Snappy;
use PhpOffice\PhpPresentation\Shape\Drawing\Base64;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide;
use ZipArchive;

/**
 * Keynote reader : the text and the images of an Apple Keynote presentation.
 *
 * A `.key` file is a Zip archive of one of two shapes, and both are read here :
 *
 * - Keynote '13 and later store the presentation as `Index/*.iwa` components : Snappy compressed
 *   frames of Protobuf messages. One component holds one slide, and the text of that slide is
 *   read out of its `TSWP.StorageArchive` messages.
 * - Keynote '09 stores it as an `index.apxl` XML document, which is read with XPath.
 *
 * Only what the format request asks for is read : the text of every slide, the speaker note and
 * the images the slide references. Anything else -- geometry beyond what an image needs, styles,
 * tables, charts, transitions -- is left to a later pass, and a `.key` of neither shape is
 * refused with an exception naming what was found.
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
     * `TSWP.StorageArchive`, the message a run of text lives in.
     *
     * @var int
     */
    protected const IWA_TYPE_TEXT_STORAGE = 2001;

    /**
     * `TSWP.StorageArchive.kind`.
     *
     * @var int
     */
    protected const IWA_FIELD_KIND = 1;

    /**
     * `TSWP.StorageArchive.text`.
     *
     * @var int
     */
    protected const IWA_FIELD_TEXT = 3;

    /**
     * The kind a speaker note is stored under : `TSWP.StorageArchive.Kind.NOTE`.
     *
     * @var int
     */
    protected const IWA_KIND_NOTE = 4;

    /**
     * How deep a message is followed while looking for the name of a data file.
     *
     * @var int
     */
    protected const IWA_MAX_DEPTH = 4;

    /**
     * Output Object.
     *
     * @var PhpPresentation
     */
    protected $oPhpPresentation;

    /**
     * The Keynote package.
     *
     * @var ZipArchive
     */
    protected $oZip;

    /**
     * Where the IWA components are : the package itself, or the `Index.zip` inside it.
     *
     * @var ZipArchive
     */
    protected $oIndex;

    /**
     * The copy of `Index.zip` on disk, if the package has one.
     *
     * @var null|string
     */
    protected $indexTempFile;

    /**
     * Every file of the `Data` folder, by lowercased name.
     *
     * @var array<string, string>
     */
    protected $arrayDataFiles = [];

    /**
     * Should the images be loaded?
     *
     * @var bool
     */
    protected $loadImages = true;

    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\Keynote read the file?
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
        if (true !== $oZip->open($pFilename)) {
            return false;
        }
        $isKeynote = null !== $this->findApxlEntry($oZip)
            || [] !== $this->findEntries($oZip, '#(^|/)Index/[^/]+\.iwa$#i')
            || null !== $this->findIndexZipEntry($oZip);
        $oZip->close();

        return $isKeynote;
    }

    /**
     * Loads PhpPresentation from file.
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
     * Reads the package and fills a presentation with what it holds.
     */
    protected function loadFile(string $pFilename): PhpPresentation
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();

        $this->oZip = new ZipArchive();
        if (true !== $this->oZip->open($pFilename)) {
            throw new InvalidFileFormatException($pFilename, self::class, 'the file is not a Zip archive');
        }

        try {
            $apxlEntry = $this->findApxlEntry($this->oZip);
            if (null !== $apxlEntry) {
                $this->loadApxl($pFilename, $apxlEntry);
            } else {
                $this->loadIwa($pFilename);
            }
        } finally {
            $this->closePackage();
        }

        return $this->oPhpPresentation;
    }

    /**
     * Closes the package and removes the copy of `Index.zip`, if there is one.
     */
    protected function closePackage(): void
    {
        if ($this->oIndex instanceof ZipArchive && $this->oIndex !== $this->oZip) {
            $this->oIndex->close();
        }
        $this->oZip->close();
        if (null !== $this->indexTempFile) {
            @unlink($this->indexTempFile);
            $this->indexTempFile = null;
        }
    }

    /**
     * The `index.apxl` document of a Keynote '09 package, compressed or not.
     */
    protected function findApxlEntry(ZipArchive $oZip): ?string
    {
        return $this->findEntries($oZip, '#(^|/)index\.apxl(\.gz)?$#i')[0] ?? null;
    }

    /**
     * The `Index.zip` some Keynote versions put the IWA components in.
     */
    protected function findIndexZipEntry(ZipArchive $oZip): ?string
    {
        return $this->findEntries($oZip, '#(^|/)Index\.zip$#i')[0] ?? null;
    }

    /**
     * Every entry of an archive whose name the pattern matches, in the order the archive lists them.
     *
     * @return array<int, string>
     */
    protected function findEntries(ZipArchive $oZip, string $pattern): array
    {
        $entries = [];
        for ($index = 0; $index < $oZip->numFiles; ++$index) {
            $name = $oZip->getNameIndex($index);
            if (is_string($name) && 1 === preg_match($pattern, $name)) {
                $entries[] = $name;
            }
        }

        return $entries;
    }

    // ------------------------------------------------------------------ IWA

    /**
     * Reads a Keynote '13 (and later) package : one slide per `Index/Slide*.iwa` component.
     */
    protected function loadIwa(string $pFilename): void
    {
        $this->oIndex = $this->openIndex($pFilename);
        $components = $this->findSlideComponents();
        if ([] === $components) {
            throw new InvalidFileFormatException($pFilename, self::class, 'no slide was found, neither as index.apxl nor as an Index/Slide*.iwa component');
        }

        foreach ($this->findEntries($this->oZip, '#(^|/)Data/[^/]+$#i') as $entry) {
            $this->arrayDataFiles[strtolower(basename($entry))] = $entry;
        }

        foreach ($components as $component) {
            $this->loadIwaSlide($component);
        }
    }

    /**
     * The archive the IWA components are read from : `Index.zip` when the package has one, and the
     * package itself otherwise.
     */
    protected function openIndex(string $pFilename): ZipArchive
    {
        $indexZipEntry = $this->findIndexZipEntry($this->oZip);
        if (null === $indexZipEntry) {
            return $this->oZip;
        }

        $contents = $this->oZip->getFromName($indexZipEntry);
        $tempFile = tempnam(sys_get_temp_dir(), 'phppresentation-keynote');
        if (false === $contents || false === $tempFile || false === file_put_contents($tempFile, $contents)) {
            throw new InvalidFileFormatException($pFilename, self::class, 'the Index.zip of the package could not be read');
        }
        $this->indexTempFile = $tempFile;

        $oIndex = new ZipArchive();
        if (true !== $oIndex->open($tempFile)) {
            throw new InvalidFileFormatException($pFilename, self::class, 'the Index.zip of the package is not a Zip archive');
        }

        return $oIndex;
    }

    /**
     * The components which hold a slide, ordered by the number Keynote names them with. A deck
     * reordered after it was written keeps the order it was written in : the tree which states the
     * order of the slides is not read yet.
     *
     * @return array<int, string>
     */
    protected function findSlideComponents(): array
    {
        $components = [];
        foreach ($this->findEntries($this->oIndex, '#(^|/)Index/Slide(-[0-9]+)?\.iwa$#i') as $entry) {
            preg_match('#Slide-([0-9]+)\.iwa$#i', $entry, $matches);
            $components[$entry] = (int) ($matches[1] ?? 0);
        }
        asort($components);

        return array_keys($components);
    }

    /**
     * Reads one slide out of the component it is stored in.
     */
    protected function loadIwaSlide(string $component): void
    {
        $contents = $this->oIndex->getFromName($component);
        if (false === $contents) {
            throw new InvalidFileFormatException($component, self::class, 'the component could not be read');
        }

        $oSlide = $this->oPhpPresentation->createSlide();
        $imageNames = [];

        foreach (IwaStream::parse(Snappy::decompressFrames($contents, $component), $component) as $archive) {
            foreach ($archive['messages'] as $message) {
                $fields = Protobuf::decode($message['payload'], $component);
                if (self::IWA_TYPE_TEXT_STORAGE === $message['type']) {
                    $this->loadIwaText($oSlide, $fields);
                }
                if ($this->loadImages) {
                    $imageNames = array_merge($imageNames, $this->findIwaImageNames($fields));
                }
            }
        }

        foreach (array_unique($imageNames) as $imageName) {
            $this->loadImage($oSlide, $this->arrayDataFiles[$imageName]);
        }
    }

    /**
     * Adds the text of a `TSWP.StorageArchive` to the slide, or to its note when that is the kind
     * of storage it is.
     *
     * @param array<int, array<int, int|string>> $fields
     */
    protected function loadIwaText(Slide $oSlide, array $fields): void
    {
        $text = implode('', Protobuf::getStrings($fields, self::IWA_FIELD_TEXT));
        if ('' === trim($text)) {
            return;
        }

        $isNote = self::IWA_KIND_NOTE === Protobuf::getInt($fields, self::IWA_FIELD_KIND);
        $oRichText = $isNote
            ? $oSlide->getNote()->createRichTextShape()
            : $oSlide->createRichTextShape();

        $this->fillRichText($oRichText, preg_split('#\r\n|\r|\n|\x{2028}#u', $text) ?: []);
    }

    /**
     * The data files a message names, at any depth. A field is taken for the name of a data file
     * only when the `Data` folder of the package holds a file of exactly that name, which is what
     * keeps a style name or a digest out of the result.
     *
     * @param array<int, array<int, int|string>> $fields
     *
     * @return array<int, string>
     */
    protected function findIwaImageNames(array $fields, int $depth = 0): array
    {
        $names = [];

        foreach ($fields as $values) {
            foreach ($values as $value) {
                if (!is_string($value)) {
                    continue;
                }
                $name = strtolower($value);
                if (isset($this->arrayDataFiles[$name])) {
                    $names[] = $name;

                    continue;
                }
                $nested = $depth < self::IWA_MAX_DEPTH ? Protobuf::tryDecode($value) : null;
                if (null !== $nested) {
                    $names = array_merge($names, $this->findIwaImageNames($nested, $depth + 1));
                }
            }
        }

        return $names;
    }

    // ----------------------------------------------------------------- APXL

    /**
     * Reads a Keynote '09 package out of its `index.apxl` document.
     */
    protected function loadApxl(string $pFilename, string $entry): void
    {
        $contents = $this->oZip->getFromName($entry);
        if (false === $contents) {
            throw new InvalidFileFormatException($pFilename, self::class, sprintf('%s could not be read', $entry));
        }
        if ('.gz' === substr(strtolower($entry), -3)) {
            $contents = @gzdecode($contents);
            if (false === $contents) {
                throw new InvalidFileFormatException($pFilename, self::class, sprintf('%s is not gzip compressed', $entry));
            }
        }

        $oXMLReader = new XMLReader();
        $oXMLReader->getDomFromString($contents);
        $oXMLReader->registerNamespace('key', self::NS_KEY);
        $oXMLReader->registerNamespace('sf', self::NS_SF);
        $oXMLReader->registerNamespace('sfa', self::NS_SFA);

        foreach ($oXMLReader->getElements('/key:presentation/key:slide-list/key:slide') as $oNodeSlide) {
            if (!$oNodeSlide instanceof DOMElement) {
                continue;
            }
            $oSlide = $this->oPhpPresentation->createSlide();
            foreach ($oXMLReader->getElements('./key:drawables/*', $oNodeSlide) as $oNodeDrawable) {
                if ($oNodeDrawable instanceof DOMElement) {
                    $this->loadApxlDrawable($oXMLReader, $oNodeDrawable, $oSlide);
                }
            }
            $this->loadApxlNote($oXMLReader, $oNodeSlide, $oSlide);
        }
    }

    /**
     * Adds a drawable of a slide, as text or as an image, to the slide.
     */
    protected function loadApxlDrawable(XMLReader $oXMLReader, DOMElement $oNodeDrawable, Slide $oSlide): void
    {
        $oNodeData = $oXMLReader->getElement('.//sf:data', $oNodeDrawable);
        if ($oNodeData instanceof DOMElement) {
            if ($this->loadImages) {
                $this->loadApxlImage($oXMLReader, $oNodeDrawable, $oNodeData, $oSlide);
            }

            return;
        }

        $paragraphs = $this->loadApxlParagraphs($oXMLReader, $oNodeDrawable);
        if ([] === $paragraphs) {
            return;
        }

        $oRichText = $oSlide->createRichTextShape();
        $this->loadApxlGeometry($oXMLReader, $oNodeDrawable, $oRichText);
        $this->fillRichText($oRichText, $paragraphs);
    }

    /**
     * Adds the image a drawable points at to the slide.
     */
    protected function loadApxlImage(XMLReader $oXMLReader, DOMElement $oNodeDrawable, DOMElement $oNodeData, Slide $oSlide): void
    {
        $entry = $this->findApxlDataEntry($oNodeData->getAttributeNS(self::NS_SF, 'path'));
        if (null === $entry) {
            return;
        }

        $oDrawing = $this->loadImage($oSlide, $entry);
        if (null !== $oDrawing) {
            $this->loadApxlGeometry($oXMLReader, $oNodeDrawable, $oDrawing);
        }
    }

    /**
     * The entry of the package a `sf:path` points at : Keynote writes it relative to the package,
     * so a name alone is looked up wherever the package keeps it.
     */
    protected function findApxlDataEntry(string $path): ?string
    {
        if ('' === $path) {
            return null;
        }
        if (is_array($this->oZip->statName($path))) {
            return $path;
        }

        return $this->findEntries($this->oZip, '#(^|/)' . preg_quote(basename($path), '#') . '$#')[0] ?? null;
    }

    /**
     * The paragraphs a text container holds, as plain text.
     *
     * @return array<int, string>
     */
    protected function loadApxlParagraphs(XMLReader $oXMLReader, DOMElement $oNodeContainer): array
    {
        $paragraphs = [];
        foreach ($oXMLReader->getElements('.//sf:text-body/sf:p', $oNodeContainer) as $oNodeParagraph) {
            if (!$oNodeParagraph instanceof DOMElement) {
                continue;
            }
            $spans = [];
            foreach ($oXMLReader->getElements('.//sf:span', $oNodeParagraph) as $oNodeSpan) {
                if ($oNodeSpan instanceof DOMElement) {
                    $spans[] = (string) $oNodeSpan->nodeValue;
                }
            }
            $paragraphs[] = [] === $spans ? (string) $oNodeParagraph->nodeValue : implode('', $spans);
        }

        return '' === trim(implode('', $paragraphs)) ? [] : $paragraphs;
    }

    /**
     * Reads the position and the size of a drawable.
     *
     * @param Base64|RichText $oShape
     */
    protected function loadApxlGeometry(XMLReader $oXMLReader, DOMElement $oNodeDrawable, $oShape): void
    {
        $oNodePosition = $oXMLReader->getElement('./sf:geometry/sf:position', $oNodeDrawable);
        if ($oNodePosition instanceof DOMElement) {
            $oShape->setOffsetX((int) $oNodePosition->getAttributeNS(self::NS_SFA, 'x'));
            $oShape->setOffsetY((int) $oNodePosition->getAttributeNS(self::NS_SFA, 'y'));
        }

        $oNodeSize = $oXMLReader->getElement('./sf:geometry/sf:size', $oNodeDrawable);
        if ($oNodeSize instanceof DOMElement) {
            $oShape->setWidth((int) $oNodeSize->getAttributeNS(self::NS_SFA, 'w'));
            $oShape->setHeight((int) $oNodeSize->getAttributeNS(self::NS_SFA, 'h'));
        }
    }

    /**
     * Reads the speaker note of a slide.
     */
    protected function loadApxlNote(XMLReader $oXMLReader, DOMElement $oNodeSlide, Slide $oSlide): void
    {
        $oNodeNote = $oXMLReader->getElement('./key:notes', $oNodeSlide);
        if (!$oNodeNote instanceof DOMElement) {
            return;
        }

        $paragraphs = $this->loadApxlParagraphs($oXMLReader, $oNodeNote);
        if ([] !== $paragraphs) {
            $this->fillRichText($oSlide->getNote()->createRichTextShape(), $paragraphs);
        }
    }

    // ---------------------------------------------------------------- Shared

    /**
     * Adds an image of the package to a slide.
     */
    protected function loadImage(Slide $oSlide, string $entry): ?Base64
    {
        $contents = $this->oZip->getFromName($entry);
        if (false === $contents || '' === $contents) {
            return null;
        }
        $size = @getimagesizefromstring($contents);
        if (false === $size) {
            return null;
        }

        $oDrawing = new Base64();
        $oDrawing->setData('data:' . $size['mime'] . ';base64,' . base64_encode($contents));
        $oDrawing->setName(basename($entry));
        $oDrawing->getShadow()->setVisible(false);
        $oDrawing->setResizeProportional(false);
        $oDrawing->setWidth((int) $size[0]);
        $oDrawing->setHeight((int) $size[1]);
        $oSlide->addShape($oDrawing);

        return $oDrawing;
    }

    /**
     * Writes plain paragraphs into a text shape.
     *
     * @param array<int, string> $paragraphs
     */
    protected function fillRichText(RichText $oRichText, array $paragraphs): void
    {
        $isFirst = true;
        foreach ($paragraphs as $paragraph) {
            $oParagraph = $isFirst ? $oRichText->getActiveParagraph() : $oRichText->createParagraph();
            $isFirst = false;
            $oParagraph->createTextRun($paragraph);
        }
    }
}
