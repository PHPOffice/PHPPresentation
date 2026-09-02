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

use DateTime;
use DOMElement;
use DOMNode;
use DOMNodeList;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLReader;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\Exception\FeatureNotImplementedException;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing\Base64;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Shape\Table\Cell;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\AbstractSlide;
use PhpOffice\PhpPresentation\Slide\Note;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Borders;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Style\SchemeColor;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Style\TextStyle;
use ZipArchive;

/**
 * Serialized format reader.
 */
class PowerPoint2007 implements ReaderInterface
{
    /**
     * The nine plots the Writer knows, by the element each is written as.
     *
     * @var array<string, class-string<Chart\Type\AbstractType>>
     */
    private const CHART_TYPES = [
        'c:areaChart' => Chart\Type\Area::class,
        'c:barChart' => Chart\Type\Bar::class,
        'c:bar3DChart' => Chart\Type\Bar3D::class,
        'c:doughnutChart' => Chart\Type\Doughnut::class,
        'c:lineChart' => Chart\Type\Line::class,
        'c:pieChart' => Chart\Type\Pie::class,
        'c:pie3DChart' => Chart\Type\Pie3D::class,
        'c:radarChart' => Chart\Type\Radar::class,
        'c:scatterChart' => Chart\Type\Scatter::class,
    ];

    /**
     * Output Object.
     *
     * @var PhpPresentation
     */
    protected $oPhpPresentation;

    /**
     * Output Object.
     *
     * @var ZipArchive
     */
    protected $oZip;

    /**
     * @var array<string, array<string, array<string, string>>>
     */
    protected $arrayRels = [];

    /**
     * @var SlideLayout[]
     */
    protected $arraySlideLayouts = [];

    /**
     * @var string
     */
    protected $filename;

    /**
     * @var string
     */
    protected $fileRels;

    /**
     * @var bool
     */
    protected $loadImages = true;

    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     */
    public function canRead(string $pFilename): bool
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     */
    public function fileSupportsUnserializePhpPresentation(string $pFilename = ''): bool
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new FileNotFoundException($pFilename);
        }

        $oZip = new ZipArchive();
        // Is it a zip ?
        if (true === $oZip->open($pFilename)) {
            // Is it an OpenXML Document ?
            // Is it a Presentation ?
            if (is_array($oZip->statName('[Content_Types].xml')) && is_array($oZip->statName('ppt/presentation.xml'))) {
                return true;
            }
        }

        return false;
    }

    /**
     * Loads PhpPresentation Serialized file.
     */
    public function load(string $pFilename, int $flags = 0): PhpPresentation
    {
        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new InvalidFileFormatException($pFilename, self::class);
        }

        $this->loadImages = !((bool) ($flags & self::SKIP_IMAGES));

        return $this->loadFile($pFilename);
    }

    /**
     * Load PhpPresentation Serialized file.
     */
    protected function loadFile(string $pFilename): PhpPresentation
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();
        $this->oPhpPresentation->setAllMasterSlides([]);
        $this->filename = $pFilename;

        $this->oZip = new ZipArchive();
        $this->oZip->open($this->filename);
        $docPropsCore = $this->oZip->getFromName('docProps/core.xml');
        if (false !== $docPropsCore) {
            $this->loadDocumentProperties($docPropsCore);
        }

        $docThumbnail = $this->oZip->getFromName('_rels/.rels');
        if ($docThumbnail !== false) {
            $this->loadThumbnailProperties($docThumbnail);
        }

        $docPropsCustom = $this->oZip->getFromName('docProps/custom.xml');
        if (false !== $docPropsCustom) {
            $this->loadCustomProperties($docPropsCustom);
        }

        $pptViewProps = $this->oZip->getFromName('ppt/viewProps.xml');
        if (false !== $pptViewProps) {
            $this->loadViewProperties($pptViewProps);
        }

        $pptPresentation = $this->oZip->getFromName('ppt/presentation.xml');
        if (false !== $pptPresentation) {
            $this->loadDocumentLayout($pptPresentation);
            $this->loadSlides($pptPresentation);
        }

        $pptPresProps = $this->oZip->getFromName('ppt/presProps.xml');
        if (false !== $pptPresProps) {
            $this->loadPresentationProperties($pptPresentation);
        }

        return $this->oPhpPresentation;
    }

    /**
     * Read Document Layout.
     */
    protected function loadDocumentLayout(string $sPart): void
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            foreach ($xmlReader->getElements('/p:presentation/p:sldSz') as $oElement) {
                if (!($oElement instanceof DOMElement)) {
                    continue;
                }
                $type = $oElement->getAttribute('type');
                $oLayout = $this->oPhpPresentation->getLayout();
                // ST_SlideSizeType allows an explicit "custom" value, which the library stores as LAYOUT_CUSTOM (an empty string)
                if (DocumentLayout::LAYOUT_CUSTOM == $type || 'custom' === $type) {
                    $oLayout->setCX((float) $oElement->getAttribute('cx'));
                    $oLayout->setCY((float) $oElement->getAttribute('cy'));
                } else {
                    $oLayout->setDocumentLayout($type, true);
                    if ($oElement->getAttribute('cx') < $oElement->getAttribute('cy')) {
                        $oLayout->setDocumentLayout($type, false);
                    }
                }
            }
        }
    }

    /**
     * Read Document Properties.
     */
    protected function loadDocumentProperties(string $sPart): void
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            $arrayProperties = [
                '/cp:coreProperties/dc:creator' => 'setCreator',
                '/cp:coreProperties/cp:lastModifiedBy' => 'setLastModifiedBy',
                '/cp:coreProperties/dc:title' => 'setTitle',
                '/cp:coreProperties/dc:description' => 'setDescription',
                '/cp:coreProperties/dc:subject' => 'setSubject',
                '/cp:coreProperties/cp:keywords' => 'setKeywords',
                '/cp:coreProperties/cp:category' => 'setCategory',
                '/cp:coreProperties/cp:revision' => 'setRevision',
                '/cp:coreProperties/cp:contentStatus' => 'setStatus',
            ];
            $arrayDateProperties = [
                '/cp:coreProperties/dcterms:created' => 'setCreated',
                '/cp:coreProperties/dcterms:modified' => 'setModified',
            ];
            $oProperties = $this->oPhpPresentation->getDocumentProperties();
            foreach ($arrayProperties as $path => $property) {
                $oElement = $xmlReader->getElement($path);
                if ($oElement instanceof DOMElement) {
                    $oProperties->{$property}((string) $oElement->nodeValue);
                }
            }
            foreach ($arrayDateProperties as $path => $property) {
                $oElement = $xmlReader->getElement($path);
                if ($oElement instanceof DOMElement) {
                    $dateTime = DateTime::createFromFormat(DateTime::W3C, (string) $oElement->nodeValue);
                    if (false !== $dateTime) {
                        $oProperties->{$property}($dateTime->getTimestamp());
                    }
                }
            }
        }
    }

    /**
     * Read information of the document thumbnail.
     */
    protected function loadThumbnailProperties(string $sPart): void
    {
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromString($sPart);

        $oElement = $xmlReader->getElement('*[@Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"]');
        if ($oElement instanceof DOMElement) {
            $path = $oElement->getAttribute('Target');
            $this->oPhpPresentation
                ->getPresentationProperties()
                ->setThumbnailPath('', PresentationProperties::THUMBNAIL_DATA, $this->oZip->getFromName($path));
        }
    }

    /**
     * Read Custom Properties.
     */
    protected function loadCustomProperties(string $sPart): void
    {
        $xmlReader = new XMLReader();
        $sPart = str_replace(' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"', '', $sPart);
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            foreach ($xmlReader->getElements('/Properties/property[@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"]') as $element) {
                if (!$element->hasAttribute('name')) {
                    continue;
                }
                $propertyName = $element->getAttribute('name');
                if ($propertyName == '_MarkAsFinal') {
                    $attributeElement = $xmlReader->getElement('vt:bool', $element);
                    if ($attributeElement && 'true' == $attributeElement->nodeValue) {
                        $this->oPhpPresentation->getPresentationProperties()->markAsFinal(true);
                    }
                } else {
                    $attributeTypeInt = $xmlReader->getElement('vt:i4', $element);
                    $attributeTypeFloat = $xmlReader->getElement('vt:r8', $element);
                    $attributeTypeBoolean = $xmlReader->getElement('vt:bool', $element);
                    $attributeTypeDate = $xmlReader->getElement('vt:filetime', $element);
                    $attributeTypeString = $xmlReader->getElement('vt:lpwstr', $element);

                    if ($attributeTypeInt) {
                        $propertyType = DocumentProperties::PROPERTY_TYPE_INTEGER;
                        $propertyValue = (int) $attributeTypeInt->nodeValue;
                    } elseif ($attributeTypeFloat) {
                        $propertyType = DocumentProperties::PROPERTY_TYPE_FLOAT;
                        $propertyValue = (float) $attributeTypeFloat->nodeValue;
                    } elseif ($attributeTypeBoolean) {
                        $propertyType = DocumentProperties::PROPERTY_TYPE_BOOLEAN;
                        $propertyValue = $attributeTypeBoolean->nodeValue == 'true' ? true : false;
                    } elseif ($attributeTypeDate) {
                        $propertyType = DocumentProperties::PROPERTY_TYPE_DATE;
                        $propertyValue = strtotime($attributeTypeDate->nodeValue);
                    } else {
                        $propertyType = DocumentProperties::PROPERTY_TYPE_STRING;
                        $propertyValue = $attributeTypeString->nodeValue;
                    }

                    $this->oPhpPresentation->getDocumentProperties()->setCustomProperty($propertyName, $propertyValue, $propertyType);
                }
            }
        }
    }

    /**
     * Read Presentation Properties.
     */
    protected function loadPresentationProperties(string $sPart): void
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            $element = $xmlReader->getElement('/p:presentationPr/p:showPr');
            if ($element instanceof DOMElement) {
                if ($element->hasAttribute('loop')) {
                    $this->oPhpPresentation->getPresentationProperties()->setLoopContinuouslyUntilEsc(
                        (bool) $element->getAttribute('loop')
                    );
                }
                if (null !== $xmlReader->getElement('p:present', $element)) {
                    $this->oPhpPresentation->getPresentationProperties()->setSlideshowType(
                        PresentationProperties::SLIDESHOW_TYPE_PRESENT
                    );
                }
                if (null !== $xmlReader->getElement('p:browse', $element)) {
                    $this->oPhpPresentation->getPresentationProperties()->setSlideshowType(
                        PresentationProperties::SLIDESHOW_TYPE_BROWSE
                    );
                }
                if (null !== $xmlReader->getElement('p:kiosk', $element)) {
                    $this->oPhpPresentation->getPresentationProperties()->setSlideshowType(
                        PresentationProperties::SLIDESHOW_TYPE_KIOSK
                    );
                }
            }
        }
    }

    /**
     * Read View Properties.
     */
    protected function loadViewProperties(string $sPart): void
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            $pathZoom = '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx';
            $oElement = $xmlReader->getElement($pathZoom);
            if ($oElement instanceof DOMElement) {
                if ($oElement->hasAttribute('d') && $oElement->hasAttribute('n')) {
                    $this->oPhpPresentation->getPresentationProperties()->setZoom((int) $oElement->getAttribute('n') / (int) $oElement->getAttribute('d'));
                }
            }
        }
    }

    /**
     * Extract all slides.
     */
    protected function loadSlides(string $sPart): void
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            $fileRels = 'ppt/_rels/presentation.xml.rels';
            $this->loadRels($fileRels);
            // Load the Masterslides
            $this->loadMasterSlides($xmlReader, $fileRels);
            // Continue with loading the slides
            foreach ($xmlReader->getElements('/p:presentation/p:sldIdLst/p:sldId') as $oElement) {
                if (!($oElement instanceof DOMElement)) {
                    continue;
                }
                $rId = $oElement->getAttribute('r:id');
                $pathSlide = isset($this->arrayRels[$fileRels][$rId]) ? $this->arrayRels[$fileRels][$rId]['Target'] : '';
                if (!empty($pathSlide)) {
                    $pptSlide = $this->oZip->getFromName('ppt/' . $pathSlide);
                    if (false !== $pptSlide) {
                        $slideRels = 'ppt/slides/_rels/' . basename($pathSlide) . '.rels';
                        $this->loadRels($slideRels);
                        $this->loadSlide($pptSlide, basename($pathSlide));
                        foreach ($this->arrayRels[$slideRels] as $rel) {
                            if ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide' == $rel['Type']) {
                                $this->loadSlideNote(basename($rel['Target']), $this->oPhpPresentation->getActiveSlide());
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Extract all MasterSlides.
     */
    protected function loadMasterSlides(XMLReader $xmlReader, string $fileRels): void
    {
        // Get all the MasterSlide Id's from the presentation.xml file
        foreach ($xmlReader->getElements('/p:presentation/p:sldMasterIdLst/p:sldMasterId') as $oElement) {
            if (!($oElement instanceof DOMElement)) {
                continue;
            }
            $rId = $oElement->getAttribute('r:id');
            // Get the path to the masterslide from the array with _rels files
            $pathMasterSlide = isset($this->arrayRels[$fileRels][$rId]) ?
                $this->arrayRels[$fileRels][$rId]['Target'] : '';
            if (!empty($pathMasterSlide)) {
                $pptMasterSlide = $this->oZip->getFromName('ppt/' . $pathMasterSlide);
                if (false !== $pptMasterSlide) {
                    $this->loadRels('ppt/slideMasters/_rels/' . basename($pathMasterSlide) . '.rels');
                    $this->loadMasterSlide($pptMasterSlide, basename($pathMasterSlide));
                }
            }
        }
    }

    /**
     * Extract data from slide.
     */
    protected function loadSlide(string $sPart, string $baseFile): void
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            $xmlReader->registerNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
            // Core
            $oSlide = $this->oPhpPresentation->createSlide();
            $this->oPhpPresentation->setActiveSlideIndex($this->oPhpPresentation->getSlideCount() - 1);
            $oSlide->setRelsIndex('ppt/slides/_rels/' . $baseFile . '.rels');

            // Name
            $oElement = $xmlReader->getElement('/p:sld/p:cSld');
            if ($oElement instanceof DOMElement && $oElement->hasAttribute('name')) {
                $oSlide->setName($oElement->getAttribute('name'));
            }

            // Background
            $oElement = $xmlReader->getElement('/p:sld/p:cSld/p:bg/p:bgPr');
            if ($oElement instanceof DOMElement) {
                $oElementColor = $xmlReader->getElement('a:solidFill/a:srgbClr', $oElement);
                if ($oElementColor instanceof DOMElement) {
                    // Color
                    $oColor = new Color();
                    $oColor->setRGB($oElementColor->hasAttribute('val') ? $oElementColor->getAttribute('val') : null);
                    // Background
                    $oBackground = new Slide\Background\Color();
                    $oBackground->setColor($oColor);
                    // Slide Background
                    $oSlide = $this->oPhpPresentation->getActiveSlide();
                    $oSlide->setBackground($oBackground);
                }
                $oElementColor = $xmlReader->getElement('a:solidFill/a:schemeClr', $oElement);
                if ($oElementColor instanceof DOMElement) {
                    // Color
                    $oColor = new SchemeColor();
                    $oColor->setValue($oElementColor->hasAttribute('val') ? $oElementColor->getAttribute('val') : null);
                    // Background
                    $oBackground = new Slide\Background\SchemeColor();
                    $oBackground->setSchemeColor($oColor);
                    // Slide Background
                    $oSlide = $this->oPhpPresentation->getActiveSlide();
                    $oSlide->setBackground($oBackground);
                }
                $oElementImage = $xmlReader->getElement('a:blipFill/a:blip', $oElement);
                if ($oElementImage instanceof DOMElement) {
                    $relImg = $this->arrayRels['ppt/slides/_rels/' . $baseFile . '.rels'][$oElementImage->getAttribute('r:embed')];
                    if (is_array($relImg)) {
                        // File
                        $pathImage = 'ppt/slides/' . $relImg['Target'];
                        $pathImage = explode('/', $pathImage);
                        foreach ($pathImage as $key => $partPath) {
                            if ('..' == $partPath) {
                                unset($pathImage[$key - 1], $pathImage[$key]);
                            }
                        }
                        $pathImage = implode('/', $pathImage);
                        $contentImg = $this->oZip->getFromName($pathImage);

                        $tmpBkgImg = tempnam(sys_get_temp_dir(), 'PhpPresentationReaderPpt2007Bkg');
                        file_put_contents($tmpBkgImg, $contentImg);
                        // Background
                        $oBackground = new Slide\Background\Image();
                        $oBackground
                            ->setPath($tmpBkgImg)
                            ->setExtension(pathinfo($pathImage, PATHINFO_EXTENSION));
                        // Slide Background
                        $oSlide = $this->oPhpPresentation->getActiveSlide();
                        $oSlide->setBackground($oBackground);
                    }
                }
            }

            // Shapes
            $arrayElements = $xmlReader->getElements('/p:sld/p:cSld/p:spTree/*');
            $this->loadSlideShapes($xmlReader, $oSlide, $arrayElements, $xmlReader);

            // Layout
            $oSlide = $this->oPhpPresentation->getActiveSlide();
            foreach ($this->arrayRels['ppt/slides/_rels/' . $baseFile . '.rels'] as $valueRel) {
                if ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' == $valueRel['Type']) {
                    $layoutBasename = basename($valueRel['Target']);
                    if (array_key_exists($layoutBasename, $this->arraySlideLayouts)) {
                        $oSlide->setSlideLayout($this->arraySlideLayouts[$layoutBasename]);
                    }

                    break;
                }
            }
        }
    }

    protected function loadMasterSlide(string $sPart, string $baseFile): void
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            // Core
            $oSlideMaster = $this->oPhpPresentation->createMasterSlide();
            $oSlideMaster->setTextStyles(new TextStyle(false));
            $oSlideMaster->setRelsIndex('ppt/slideMasters/_rels/' . $baseFile . '.rels');

            // Background
            $oElement = $xmlReader->getElement('/p:sldMaster/p:cSld/p:bg');
            if ($oElement instanceof DOMElement) {
                $this->loadSlideBackground($xmlReader, $oElement, $oSlideMaster);
            }

            // Shapes
            $arrayElements = $xmlReader->getElements('/p:sldMaster/p:cSld/p:spTree/*');
            $this->loadSlideShapes($xmlReader, $oSlideMaster, $arrayElements, $xmlReader);

            // Header & Footer

            // ColorMapping
            $colorMap = [];
            $oElement = $xmlReader->getElement('/p:sldMaster/p:clrMap');
            if ($oElement->hasAttributes()) {
                foreach ($oElement->attributes as $attr) {
                    $colorMap[$attr->nodeName] = $attr->nodeValue;
                }
                $oSlideMaster->colorMap->setMapping($colorMap);
            }

            // TextStyles
            $arrayElementTxStyles = $xmlReader->getElements('/p:sldMaster/p:txStyles/*');
            foreach ($arrayElementTxStyles as $oElementTxStyle) {
                $arrayElementsLvl = $xmlReader->getElements('/p:sldMaster/p:txStyles/' . $oElementTxStyle->nodeName . '/*');
                foreach ($arrayElementsLvl as $oElementLvl) {
                    if (!($oElementLvl instanceof DOMElement) || 'a:extLst' == $oElementLvl->nodeName) {
                        continue;
                    }
                    $oRTParagraph = new Paragraph();

                    if ('a:defPPr' == $oElementLvl->nodeName) {
                        $level = 0;
                    } else {
                        $level = str_replace('a:lvl', '', $oElementLvl->nodeName);
                        $level = str_replace('pPr', '', $level);
                        $level = (int) $level;
                    }

                    if ($oElementLvl->hasAttribute('algn')) {
                        $oRTParagraph->getAlignment()->setHorizontal($oElementLvl->getAttribute('algn'));
                    }
                    if ($oElementLvl->hasAttribute('marL')) {
                        $val = (int) $oElementLvl->getAttribute('marL');
                        $val = (int) CommonDrawing::emuToPixels((int) $val);
                        $oRTParagraph->getAlignment()->setMarginLeft($val);
                    }
                    if ($oElementLvl->hasAttribute('marR')) {
                        $val = (int) $oElementLvl->getAttribute('marR');
                        $val = (int) CommonDrawing::emuToPixels((int) $val);
                        $oRTParagraph->getAlignment()->setMarginRight($val);
                    }
                    if ($oElementLvl->hasAttribute('indent')) {
                        $val = (int) $oElementLvl->getAttribute('indent');
                        $val = (int) CommonDrawing::emuToPixels((int) $val);
                        $oRTParagraph->getAlignment()->setIndent($val);
                    }
                    $oElementLvlDefRPR = $xmlReader->getElement('a:defRPr', $oElementLvl);
                    if ($oElementLvlDefRPR instanceof DOMElement) {
                        if ($oElementLvlDefRPR->hasAttribute('sz')) {
                            $oRTParagraph->getFont()->setSize((int) ((int) $oElementLvlDefRPR->getAttribute('sz') / 100));
                        }
                        if ($oElementLvlDefRPR->hasAttribute('b') && 1 == $oElementLvlDefRPR->getAttribute('b')) {
                            $oRTParagraph->getFont()->setBold(true);
                        }
                        if ($oElementLvlDefRPR->hasAttribute('i') && 1 == $oElementLvlDefRPR->getAttribute('i')) {
                            $oRTParagraph->getFont()->setItalic(true);
                        }
                    }
                    $oElementSchemeColor = $xmlReader->getElement('a:defRPr/a:solidFill/a:schemeClr', $oElementLvl);
                    if ($oElementSchemeColor instanceof DOMElement) {
                        if ($oElementSchemeColor->hasAttribute('val')) {
                            $oSchemeColor = new SchemeColor();
                            $oSchemeColor->setValue($oElementSchemeColor->getAttribute('val'));
                            $oRTParagraph->getFont()->setColor($oSchemeColor);
                        }
                    }

                    switch ($oElementTxStyle->nodeName) {
                        case 'p:bodyStyle':
                            $oSlideMaster->getTextStyles()->setBodyStyleAtLvl($oRTParagraph, $level);

                            break;
                        case 'p:otherStyle':
                            $oSlideMaster->getTextStyles()->setOtherStyleAtLvl($oRTParagraph, $level);

                            break;
                        case 'p:titleStyle':
                            $oSlideMaster->getTextStyles()->setTitleStyleAtLvl($oRTParagraph, $level);

                            break;
                    }
                }
            }

            // Load the theme
            foreach ($this->arrayRels[$oSlideMaster->getRelsIndex()] as $arrayRel) {
                if ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' == $arrayRel['Type']) {
                    $pptTheme = $this->oZip->getFromName('ppt/' . substr($arrayRel['Target'], strrpos($arrayRel['Target'], '../') + 3));
                    if (false !== $pptTheme) {
                        $this->loadTheme($pptTheme, $oSlideMaster);
                    }

                    break;
                }
            }

            // Load the Layoutslide
            foreach ($xmlReader->getElements('/p:sldMaster/p:sldLayoutIdLst/p:sldLayoutId') as $oElement) {
                if (!($oElement instanceof DOMElement)) {
                    continue;
                }
                $rId = $oElement->getAttribute('r:id');
                // Get the path to the masterslide from the array with _rels files
                $pathLayoutSlide = isset($this->arrayRels[$oSlideMaster->getRelsIndex()][$rId]) ?
                    $this->arrayRels[$oSlideMaster->getRelsIndex()][$rId]['Target'] : '';
                if (!empty($pathLayoutSlide)) {
                    $pptLayoutSlide = $this->oZip->getFromName('ppt/' . substr($pathLayoutSlide, strrpos($pathLayoutSlide, '../') + 3));
                    if (false !== $pptLayoutSlide) {
                        $this->loadRels('ppt/slideLayouts/_rels/' . basename($pathLayoutSlide) . '.rels');
                        $oSlideMaster->addSlideLayout(
                            $this->loadLayoutSlide($pptLayoutSlide, basename($pathLayoutSlide), $oSlideMaster)
                        );
                    }
                }
            }
        }
    }

    protected function loadLayoutSlide(string $sPart, string $baseFile, SlideMaster $oSlideMaster): ?SlideLayout
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            // Core
            $oSlideLayout = new SlideLayout($oSlideMaster);
            $oSlideLayout->setRelsIndex('ppt/slideLayouts/_rels/' . $baseFile . '.rels');

            // Name
            $oElement = $xmlReader->getElement('/p:sldLayout/p:cSld');
            if ($oElement instanceof DOMElement && $oElement->hasAttribute('name')) {
                $oSlideLayout->setLayoutName($oElement->getAttribute('name'));
            }

            // Background
            $oElement = $xmlReader->getElement('/p:sldLayout/p:cSld/p:bg');
            if ($oElement instanceof DOMElement) {
                $this->loadSlideBackground($xmlReader, $oElement, $oSlideLayout);
            }

            // ColorMapping
            $oElement = $xmlReader->getElement('/p:sldLayout/p:clrMapOvr/a:overrideClrMapping');
            if ($oElement instanceof DOMElement && $oElement->hasAttributes()) {
                $colorMap = [];
                foreach ($oElement->attributes as $attr) {
                    $colorMap[$attr->nodeName] = $attr->nodeValue;
                }
                $oSlideLayout->colorMap->setMapping($colorMap);
            }

            // Shapes
            $oElements = $xmlReader->getElements('/p:sldLayout/p:cSld/p:spTree/*');
            $this->loadSlideShapes($xmlReader, $oSlideLayout, $oElements, $xmlReader);
            $this->arraySlideLayouts[$baseFile] = &$oSlideLayout;

            return $oSlideLayout;
        }

        // @phpstan-ignore-next-line
        return null;
    }

    protected function loadTheme(string $sPart, SlideMaster $oSlideMaster): void
    {
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            $oElements = $xmlReader->getElements('/a:theme/a:themeElements/a:clrScheme/*');
            foreach ($oElements as $oElement) {
                if ($oElement instanceof DOMElement) {
                    $oSchemeColor = new SchemeColor();
                    $oSchemeColor->setValue(str_replace('a:', '', $oElement->tagName));
                    $colorElement = $xmlReader->getElement('*', $oElement);
                    if ($colorElement instanceof DOMElement) {
                        if ($colorElement->hasAttribute('lastClr')) {
                            $oSchemeColor->setRGB($colorElement->getAttribute('lastClr'));
                        } elseif ($colorElement->hasAttribute('val')) {
                            $oSchemeColor->setRGB($colorElement->getAttribute('val'));
                        }
                    }
                    $oSlideMaster->addSchemeColor($oSchemeColor);
                }
            }
        }
    }

    protected function loadSlideBackground(XMLReader $xmlReader, DOMElement $oElement, AbstractSlide $oSlide): void
    {
        // Background color
        $oElementColor = $xmlReader->getElement('p:bgPr/a:solidFill/a:srgbClr', $oElement);
        if ($oElementColor instanceof DOMElement) {
            // Color
            $oColor = new Color();
            $oColor->setRGB($oElementColor->hasAttribute('val') ? $oElementColor->getAttribute('val') : null);
            // Background
            $oBackground = new Slide\Background\Color();
            $oBackground->setColor($oColor);
            // Slide Background
            $oSlide->setBackground($oBackground);
        }

        // Background scheme color
        $oElementSchemeColor = $xmlReader->getElement('p:bgRef/a:schemeClr', $oElement);
        if ($oElementSchemeColor instanceof DOMElement) {
            // Color
            $oColor = new SchemeColor();
            $oColor->setValue($oElementSchemeColor->hasAttribute('val') ? $oElementSchemeColor->getAttribute('val') : null);
            // Background
            $oBackground = new Slide\Background\SchemeColor();
            $oBackground->setSchemeColor($oColor);
            // Slide Background
            $oSlide->setBackground($oBackground);
        }

        // Background image
        $oElementImage = $xmlReader->getElement('p:bgPr/a:blipFill/a:blip', $oElement);
        if ($oElementImage instanceof DOMElement) {
            $relImg = $this->arrayRels[$oSlide->getRelsIndex()][$oElementImage->getAttribute('r:embed')];
            if (is_array($relImg)) {
                // File
                $pathImage = 'ppt/slides/' . $relImg['Target'];
                $pathImage = explode('/', $pathImage);
                foreach ($pathImage as $key => $partPath) {
                    if ('..' == $partPath) {
                        unset($pathImage[$key - 1], $pathImage[$key]);
                    }
                }
                $pathImage = implode('/', $pathImage);
                $contentImg = $this->oZip->getFromName($pathImage);

                $tmpBkgImg = tempnam(sys_get_temp_dir(), 'PhpPresentationReaderPpt2007Bkg');
                file_put_contents($tmpBkgImg, $contentImg);
                // Background
                $oBackground = new Slide\Background\Image();
                $oBackground->setPath($tmpBkgImg);
                // Slide Background
                $oSlide->setBackground($oBackground);
            }
        }
    }

    protected function loadSlideNote(string $baseFile, Slide $oSlide): void
    {
        $sPart = $this->oZip->getFromName('ppt/notesSlides/' . $baseFile);
        $xmlReader = new XMLReader();
        // @phpstan-ignore-next-line
        if ($xmlReader->getDomFromString($sPart)) {
            $oNote = $oSlide->getNote();

            $arrayElements = $xmlReader->getElements('/p:notes/p:cSld/p:spTree/*');
            $this->loadSlideShapes($xmlReader, $oNote, $arrayElements, $xmlReader);
        }
    }

    /**
     * Read the decorative flag, stored as an extension of the non-visual properties of the shape.
     *
     * @param DOMElement $node the `p:cNvPr` element of the shape
     *
     * @return bool false when the shape says nothing about it
     */
    protected function loadShapeDecorative(XMLReader $document, DOMElement $node): bool
    {
        $document->registerNamespace('adec', 'http://schemas.microsoft.com/office/drawing/2017/decorative');
        $oElement = $document->getElement('a:extLst/a:ext[@uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"]/adec:decorative', $node);
        if (!$oElement instanceof DOMElement) {
            return false;
        }

        return in_array($oElement->getAttribute('val'), ['1', 'true'], true);
    }

    /**
     * Load a group of shapes.
     *
     * @param AbstractSlide|Note $oSlide
     */
    protected function loadShapeGroup(XMLReader $document, DOMElement $node, $oSlide, XMLReader $xmlReader, ShapeContainerInterface $oContainer): void
    {
        $oShape = new Group();
        $oContainer->addShape($oShape);

        $oElement = $document->getElement('p:nvGrpSpPr/p:cNvPr', $node);
        if ($oElement instanceof DOMElement) {
            $oShape->setName($oElement->hasAttribute('name') ? $oElement->getAttribute('name') : '');
            $oShape->setDescription($oElement->hasAttribute('descr') ? $oElement->getAttribute('descr') : '');
            $oShape->setDecorative($this->loadShapeDecorative($document, $oElement));
        }

        $oElement = $document->getElement('p:grpSpPr/a:xfrm', $node);
        if ($oElement instanceof DOMElement && $oElement->hasAttribute('rot')) {
            $oShape->setRotation((int) CommonDrawing::angleToDegrees((int) $oElement->getAttribute('rot')));
        }

        $this->loadSlideShapes($document, $oSlide, $node->childNodes, $xmlReader, $oShape);
        $this->loadGroupTransform($document, $node, $oShape);
    }

    /**
     * Map the shapes of a group out of the child coordinate space declared by
     * a:chOff/a:chExt and into the parent space declared by a:off/a:ext.
     *
     * A group carries no coordinates of its own in this library: its offset and
     * extent are derived from the shapes it holds. So the mapping is baked into
     * those shapes here, which leaves the group where the file says it is.
     */
    protected function loadGroupTransform(XMLReader $document, DOMElement $node, Group $oShape): void
    {
        $oXfrm = $document->getElement('p:grpSpPr/a:xfrm', $node);
        if (!$oXfrm instanceof DOMElement) {
            return;
        }
        $off = $this->loadGroupPoint($document, $oXfrm, 'a:off', 'x', 'y');
        $chOff = $this->loadGroupPoint($document, $oXfrm, 'a:chOff', 'x', 'y');
        if (null === $off || null === $chOff) {
            return;
        }
        $ext = $this->loadGroupPoint($document, $oXfrm, 'a:ext', 'cx', 'cy');
        $chExt = $this->loadGroupPoint($document, $oXfrm, 'a:chExt', 'cx', 'cy');
        // A missing or zero a:chExt means the group is not scaled.
        $scaleX = null !== $ext && null !== $chExt && 0 != $chExt[0] ? $ext[0] / $chExt[0] : 1.0;
        $scaleY = null !== $ext && null !== $chExt && 0 != $chExt[1] ? $ext[1] / $chExt[1] : 1.0;
        if ($off === $chOff && 1.0 === $scaleX && 1.0 === $scaleY) {
            return;
        }

        $this->applyGroupTransform($oShape, $off, $chOff, $scaleX, $scaleY);
    }

    /**
     * Read one a:off/a:ext/a:chOff/a:chExt pair, in pixels.
     *
     * Truncated to whole pixels, as every shape this is subtracted from already is:
     * a shape written a third of a pixel past a:chOff has lost that third by the time
     * it gets here, so a:chOff keeps the same grid and the two losses cancel. Keeping
     * the fraction on one side of the subtraction only makes the difference worse,
     * and a group scaled by eight multiplies whatever is left over.
     *
     * @return null|array{0: int, 1: int}
     */
    protected function loadGroupPoint(XMLReader $document, DOMElement $oXfrm, string $name, string $attrX, string $attrY): ?array
    {
        $oElement = $document->getElement($name, $oXfrm);
        if (!$oElement instanceof DOMElement) {
            return null;
        }

        return [
            (int) CommonDrawing::emuToPixels((int) $oElement->getAttribute($attrX)),
            (int) CommonDrawing::emuToPixels((int) $oElement->getAttribute($attrY)),
        ];
    }

    /**
     * @param array{0: int, 1: int} $off
     * @param array{0: int, 1: int} $chOff
     */
    protected function applyGroupTransform(Group $oShape, array $off, array $chOff, float $scaleX, float $scaleY): void
    {
        foreach ($oShape->getShapeCollection() as $shape) {
            if ($shape instanceof Group) {
                // A nested group has no coordinates to move: it reports those of
                // the shapes it holds, so the mapping goes straight through it.
                $this->applyGroupTransform($shape, $off, $chOff, $scaleX, $scaleY);

                continue;
            }
            $shape->setOffsetX((int) round(($shape->getOffsetX() - $chOff[0]) * $scaleX + $off[0]));
            $shape->setOffsetY((int) round(($shape->getOffsetY() - $chOff[1]) * $scaleY + $off[1]));
            if (1.0 !== $scaleX) {
                $shape->setWidth((int) round($shape->getWidth() * $scaleX));
            }
            if (1.0 !== $scaleY) {
                $shape->setHeight((int) round($shape->getHeight() * $scaleY));
            }
        }
    }

    protected function loadShapeDrawing(XMLReader $document, DOMElement $node, AbstractSlide $oSlide, ?ShapeContainerInterface $oContainer = null): void
    {
        $oContainer = $oContainer ?? $oSlide;
        // Core
        $document->registerNamespace('asvg', 'http://schemas.microsoft.com/office/drawing/2016/SVG/main');
        if ($document->getElement('p:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip', $node)) {
            $oShape = new Base64();
        } else {
            $oShape = new Gd();
        }
        $oShape->getShadow()->setVisible(false);
        // Variables
        $fileRels = $oSlide->getRelsIndex();

        $oElement = $document->getElement('p:nvPicPr/p:cNvPr', $node);
        if ($oElement instanceof DOMElement) {
            $oShape->setName($oElement->hasAttribute('name') ? $oElement->getAttribute('name') : '');
            $oShape->setDescription($oElement->hasAttribute('descr') ? $oElement->getAttribute('descr') : '');
            $oShape->setDecorative($this->loadShapeDecorative($document, $oElement));

            // Hyperlink
            $oElementHlinkClick = $document->getElement('a:hlinkClick', $oElement);
            if (is_object($oElementHlinkClick)) {
                $oShape->setHyperlink(
                    $this->loadHyperlink($document, $oElementHlinkClick, $oShape->getHyperlink())
                );
            }
        }

        $oElement = $document->getElement('p:blipFill/a:blip', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('r:embed') && isset($this->arrayRels[$fileRels][$oElement->getAttribute('r:embed')]['Target'])) {
                $pathImage = 'ppt/slides/' . $this->arrayRels[$fileRels][$oElement->getAttribute('r:embed')]['Target'];
                $pathImage = explode('/', $pathImage);
                foreach ($pathImage as $key => $partPath) {
                    if ('..' == $partPath) {
                        unset($pathImage[$key - 1], $pathImage[$key]);
                    }
                }
                $pathImage = implode('/', $pathImage);
                $imageFile = $this->oZip->getFromName($pathImage);
                if (!empty($imageFile)) {
                    if ($oShape instanceof Gd) {
                        $info = getimagesizefromstring($imageFile);
                        if (!$info) {
                            return;
                        }
                        $oShape->setMimeType($info['mime']);
                        $oShape->setRenderingFunction(str_replace('/', '', $info['mime']));
                        $image = @imagecreatefromstring($imageFile);
                        if (!$image) {
                            return;
                        }
                        $oShape->setImageResource($image);
                    } elseif ($oShape instanceof Base64) {
                        $oShape->setData('data:image/svg+xml;base64,' . base64_encode($imageFile));
                    }
                }
            }
        }

        $oElement = $document->getElement('p:spPr', $node);
        if ($oElement instanceof DOMElement) {
            $oFill = $this->loadStyleFill($document, $oElement);
            $oShape->setFill($oFill);
        }

        $oElement = $document->getElement('p:spPr/a:xfrm', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('rot')) {
                $oShape->setRotation((int) CommonDrawing::angleToDegrees((int) $oElement->getAttribute('rot')));
            }
        }

        $oElement = $document->getElement('p:spPr/a:xfrm/a:off', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('x')) {
                $oShape->setOffsetX((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('x')));
            }
            if ($oElement->hasAttribute('y')) {
                $oShape->setOffsetY((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('y')));
            }
        }

        $oElement = $document->getElement('p:spPr/a:xfrm/a:ext', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('cx')) {
                $oShape->setWidth((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('cx')));
            }
            if ($oElement->hasAttribute('cy')) {
                $oShape->setHeight((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('cy')));
            }
        }
        // Load shape effects
        $oElement = $document->getElement('p:spPr/a:effectLst', $node);
        if ($oElement instanceof DOMElement) {
            $oShape->setShadow(
                $this->loadShadow($document, $oElement)
            );
        }
        $oContainer->addShape($oShape);
    }

    /**
     * Load Shadow for shape or paragraph.
     */
    protected function loadShadow(XMLReader $document, DOMElement $node): ?Shadow
    {
        $aNodes = $document->getElements('*', $node);
        foreach ($aNodes as $nodeShadow) {
            $type = explode(':', $nodeShadow->tagName);
            $type = array_pop($type);
            if ($type == Shadow::TYPE_SHADOW_INNER || $type == Shadow::TYPE_SHADOW_OUTER || $type == Shadow::TYPE_REFLECTION) {
                $oShadow = new Shadow();
                $oShadow->setVisible(true);
                $oShadow->setType($type);
                if ($nodeShadow->hasAttribute('blurRad')) {
                    $oShadow->setBlurRadius((int) CommonDrawing::emuToPixels((int) $nodeShadow->getAttribute('blurRad')));
                }
                if ($nodeShadow->hasAttribute('dist')) {
                    $oShadow->setDistance((int) CommonDrawing::emuToPixels((int) $nodeShadow->getAttribute('dist')));
                }
                if ($nodeShadow->hasAttribute('dir')) {
                    $oShadow->setDirection((int) CommonDrawing::angleToDegrees((int) $nodeShadow->getAttribute('dir')));
                }
                if ($nodeShadow->hasAttribute('algn')) {
                    $oShadow->setAlignment($nodeShadow->getAttribute('algn'));
                }

                // The colour is written as `a:srgbClr` and only a preset one was read back.
                // `loadStyleColor()` already reads a colour and the `a:alpha` inside it, and the
                // alpha a shadow's colour carries is the shadow's own -- it is the one the Writer
                // puts there -- so both are taken from it rather than parsed a second time here.
                foreach (['a:srgbClr', 'a:prstClr'] as $colorElement) {
                    $oSubElement = $document->getElement($colorElement, $nodeShadow);
                    if (!$oSubElement instanceof DOMElement || !$oSubElement->hasAttribute('val')) {
                        continue;
                    }
                    $oColor = $this->loadStyleColor($document, $oSubElement);
                    $oShadow->setColor($oColor);
                    $oShadow->setAlpha($oColor->getAlpha());

                    break;
                }

                return $oShadow;
            }
        }

        return null;
    }

    /**
     * @param AbstractSlide|Note $oSlide
     */
    protected function loadShapeRichText(XMLReader $document, DOMElement $node, $oSlide, ?ShapeContainerInterface $oContainer = null): void
    {
        // Core
        $oShape = new RichText();
        ($oContainer ?? $oSlide)->addShape($oShape);
        $oShape->setParagraphs([]);
        // Variables
        if ($oSlide instanceof AbstractSlide) {
            $this->fileRels = $oSlide->getRelsIndex();
        }

        $oElement = $document->getElement('p:nvSpPr/p:cNvPr', $node);
        if ($oElement instanceof DOMElement) {
            $oShape->setName($oElement->hasAttribute('name') ? $oElement->getAttribute('name') : '');
            $oShape->setDescription($oElement->hasAttribute('descr') ? $oElement->getAttribute('descr') : '');
            $oShape->setDecorative($this->loadShapeDecorative($document, $oElement));
        }

        $oElement = $document->getElement('p:spPr/a:xfrm', $node);
        if ($oElement instanceof DOMElement && $oElement->hasAttribute('rot')) {
            $oShape->setRotation((int) CommonDrawing::angleToDegrees((int) $oElement->getAttribute('rot')));
        }

        $oElement = $document->getElement('p:spPr/a:xfrm/a:off', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('x')) {
                $oShape->setOffsetX((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('x')));
            }
            if ($oElement->hasAttribute('y')) {
                $oShape->setOffsetY((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('y')));
            }
        }

        $oElement = $document->getElement('p:spPr/a:xfrm/a:ext', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('cx')) {
                $oShape->setWidth((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('cx')));
            }
            if ($oElement->hasAttribute('cy')) {
                $oShape->setHeight((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('cy')));
            }
        }

        $oElement = $document->getElement('p:nvSpPr/p:nvPr/p:ph', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('type')) {
                $placeholder = new Placeholder($oElement->getAttribute('type'));
                $oShape->setPlaceHolder($placeholder);
            }
        }

        // Load shape effects
        $oElement = $document->getElement('p:spPr/a:effectLst', $node);
        if ($oElement instanceof DOMElement) {
            $oShape->setShadow(
                $this->loadShadow($document, $oElement)
            );
        }

        // FBU-20210202+ Read body definitions
        $bodyPr = $document->getElement('p:txBody/a:bodyPr', $node);
        if ($bodyPr instanceof DOMElement) {
            // the insets are EMU in the file and pixels in the model, which is the unit the
            // Writer converts back from
            if ($bodyPr->hasAttribute('lIns')) {
                $oShape->setInsetLeft(CommonDrawing::emuToPixels((int) $bodyPr->getAttribute('lIns')));
            }
            if ($bodyPr->hasAttribute('tIns')) {
                $oShape->setInsetTop(CommonDrawing::emuToPixels((int) $bodyPr->getAttribute('tIns')));
            }
            if ($bodyPr->hasAttribute('rIns')) {
                $oShape->setInsetRight(CommonDrawing::emuToPixels((int) $bodyPr->getAttribute('rIns')));
            }
            if ($bodyPr->hasAttribute('bIns')) {
                $oShape->setInsetBottom(CommonDrawing::emuToPixels((int) $bodyPr->getAttribute('bIns')));
            }
            if ($bodyPr->hasAttribute('anchorCtr')) {
                $oShape->setVerticalAlignCenter((int) $bodyPr->getAttribute('anchorCtr'));
            }
            if ($bodyPr->hasAttribute('rtlCol')) {
                $oShape->setColumnsRTL((bool) (int) $bodyPr->getAttribute('rtlCol'));
            }
            // `none` and `square` are the two values a shape's wrap holds, spelled the same way
            if ($bodyPr->hasAttribute('wrap')) {
                $oShape->setWrap($bodyPr->getAttribute('wrap'));
            }
        }

        $arrayElements = $document->getElements('p:txBody/a:p', $node);
        foreach ($arrayElements as $oElement) {
            if ($oElement instanceof DOMElement) {
                $this->loadParagraph($document, $oElement, $oShape);
            }
        }

        $oElement = $document->getElement('p:spPr', $node);
        if ($oElement instanceof DOMElement) {
            $oShape->setFill(
                $this->loadStyleFill($document, $oElement)
            );
        }

        if (count($oShape->getParagraphs()) > 0) {
            $oShape->setActiveParagraph(0);
        }
    }

    protected function loadShapeTable(XMLReader $document, DOMElement $node, AbstractSlide $oSlide, ?ShapeContainerInterface $oContainer = null): void
    {
        $this->fileRels = $oSlide->getRelsIndex();

        $oShape = new Table();
        ($oContainer ?? $oSlide)->addShape($oShape);

        $oElement = $document->getElement('p:cNvPr', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('name')) {
                $oShape->setName($oElement->getAttribute('name'));
            }
            if ($oElement->hasAttribute('descr')) {
                $oShape->setDescription($oElement->getAttribute('descr'));
            }
            $oShape->setDecorative($this->loadShapeDecorative($document, $oElement));
        }

        $oElement = $document->getElement('p:xfrm/a:off', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('x')) {
                $oShape->setOffsetX((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('x')));
            }
            if ($oElement->hasAttribute('y')) {
                $oShape->setOffsetY((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('y')));
            }
        }

        $oElement = $document->getElement('p:xfrm/a:ext', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('cx')) {
                $oShape->setWidth((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('cx')));
            }
            if ($oElement->hasAttribute('cy')) {
                $oShape->setHeight((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('cy')));
            }
        }

        $oElement = $document->getElement('a:graphic/a:graphicData/a:tbl/a:tblPr', $node);
        // Both attributes default to false when the element or the attribute is absent
        $oShape->setFirstRow($oElement instanceof DOMElement && in_array($oElement->getAttribute('firstRow'), ['1', 'true'], true));
        $oShape->setBandRow($oElement instanceof DOMElement && in_array($oElement->getAttribute('bandRow'), ['1', 'true'], true));

        $arrayElements = $document->getElements('a:graphic/a:graphicData/a:tbl/a:tblGrid/a:gridCol', $node);
        $oShape->setNumColumns($arrayElements->length);
        $oShape->createRow();
        foreach ($arrayElements as $key => $oElement) {
            if ($oElement instanceof DOMElement && $oElement->getAttribute('w')) {
                $oShape->getRow(0)->getCell($key)->setWidth((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('w')));
            }
        }

        $arrayElements = $document->getElements('a:graphic/a:graphicData/a:tbl/a:tr', $node);
        foreach ($arrayElements as $keyRow => $oElementRow) {
            if (!($oElementRow instanceof DOMElement)) {
                continue;
            }
            if ($oShape->hasRow($keyRow)) {
                $oRow = $oShape->getRow($keyRow);
            } else {
                $oRow = $oShape->createRow();
            }
            if ($oElementRow->hasAttribute('h')) {
                $oRow->setHeight((int) CommonDrawing::emuToPixels((int) $oElementRow->getAttribute('h')));
            }
            $arrayElementsCell = $document->getElements('a:tc', $oElementRow);
            foreach ($arrayElementsCell as $keyCell => $oElementCell) {
                if (!($oElementCell instanceof DOMElement)) {
                    continue;
                }
                $oCell = $oRow->getCell($keyCell);
                $oCell->setParagraphs([]);
                if ($oElementCell->hasAttribute('gridSpan')) {
                    $oCell->setColSpan((int) $oElementCell->getAttribute('gridSpan'));
                }
                if ($oElementCell->hasAttribute('rowSpan')) {
                    $oCell->setRowSpan((int) $oElementCell->getAttribute('rowSpan'));
                }

                foreach ($document->getElements('a:txBody/a:p', $oElementCell) as $oElementPara) {
                    if ($oElementPara instanceof DOMElement) {
                        $this->loadParagraph($document, $oElementPara, $oCell);
                    }
                }

                $oElementTcPr = $document->getElement('a:tcPr', $oElementCell);
                if ($oElementTcPr instanceof DOMElement) {
                    $numParagraphs = count($oCell->getParagraphs());
                    if ($numParagraphs > 0) {
                        if ($oElementTcPr->hasAttribute('vert')) {
                            $oCell->getParagraph(0)->getAlignment()->setTextDirection($oElementTcPr->getAttribute('vert'));
                        }
                        if ($oElementTcPr->hasAttribute('anchor')) {
                            $oCell->getParagraph(0)->getAlignment()->setVertical($oElementTcPr->getAttribute('anchor'));
                        }
                        if ($oElementTcPr->hasAttribute('marB')) {
                            $oCell->getParagraph(0)->getAlignment()->setMarginBottom(CommonDrawing::emuToPixels((int) $oElementTcPr->getAttribute('marB')));
                        }
                        if ($oElementTcPr->hasAttribute('marL')) {
                            $oCell->getParagraph(0)->getAlignment()->setMarginLeft(CommonDrawing::emuToPixels((int) $oElementTcPr->getAttribute('marL')));
                        }
                        if ($oElementTcPr->hasAttribute('marR')) {
                            $oCell->getParagraph(0)->getAlignment()->setMarginRight(CommonDrawing::emuToPixels((int) $oElementTcPr->getAttribute('marR')));
                        }
                        if ($oElementTcPr->hasAttribute('marT')) {
                            $oCell->getParagraph(0)->getAlignment()->setMarginTop(CommonDrawing::emuToPixels((int) $oElementTcPr->getAttribute('marT')));
                        }
                    }

                    $oFill = $this->loadStyleFill($document, $oElementTcPr);
                    if ($oFill instanceof Fill) {
                        $oCell->setFill($oFill);
                    }

                    $oBorders = new Borders();
                    $oElementBorderL = $document->getElement('a:lnL', $oElementTcPr);
                    if ($oElementBorderL instanceof DOMElement) {
                        $this->loadStyleBorder($document, $oElementBorderL, $oBorders->getLeft());
                    }
                    $oElementBorderR = $document->getElement('a:lnR', $oElementTcPr);
                    if ($oElementBorderR instanceof DOMElement) {
                        $this->loadStyleBorder($document, $oElementBorderR, $oBorders->getRight());
                    }
                    $oElementBorderT = $document->getElement('a:lnT', $oElementTcPr);
                    if ($oElementBorderT instanceof DOMElement) {
                        $this->loadStyleBorder($document, $oElementBorderT, $oBorders->getTop());
                    }
                    $oElementBorderB = $document->getElement('a:lnB', $oElementTcPr);
                    if ($oElementBorderB instanceof DOMElement) {
                        $this->loadStyleBorder($document, $oElementBorderB, $oBorders->getBottom());
                    }
                    $oElementBorderDiagDown = $document->getElement('a:lnTlToBr', $oElementTcPr);
                    if ($oElementBorderDiagDown instanceof DOMElement) {
                        $this->loadStyleBorder($document, $oElementBorderDiagDown, $oBorders->getDiagonalDown());
                    }
                    $oElementBorderDiagUp = $document->getElement('a:lnBlToTr', $oElementTcPr);
                    if ($oElementBorderDiagUp instanceof DOMElement) {
                        $this->loadStyleBorder($document, $oElementBorderDiagUp, $oBorders->getDiagonalUp());
                    }
                    $oCell->setBorders($oBorders);
                }
            }
        }
    }

    protected function loadShapeChart(XMLReader $document, DOMElement $node, AbstractSlide $oSlide, ?ShapeContainerInterface $oContainer = null): void
    {
        $oContainer = $oContainer ?? $oSlide;
        $this->fileRels = $oSlide->getRelsIndex();

        $oShape = new Chart();

        $oElement = $document->getElement('p:cNvPr', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('name')) {
                $oShape->setName($oElement->getAttribute('name'));
            }
            if ($oElement->hasAttribute('descr')) {
                $oShape->setDescription($oElement->getAttribute('descr'));
            }
            $oShape->setDecorative($this->loadShapeDecorative($document, $oElement));
        }

        $oElement = $document->getElement('p:xfrm/a:off', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('x')) {
                $oShape->setOffsetX((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('x')));
            }
            if ($oElement->hasAttribute('y')) {
                $oShape->setOffsetY((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('y')));
            }
        }

        $oElement = $document->getElement('p:xfrm/a:ext', $node);
        if ($oElement instanceof DOMElement) {
            if ($oElement->hasAttribute('cx')) {
                $oShape->setWidth((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('cx')));
            }
            if ($oElement->hasAttribute('cy')) {
                $oShape->setHeight((int) CommonDrawing::emuToPixels((int) $oElement->getAttribute('cy')));
            }
        }

        $chartElement = $document->getElement('a:graphic/a:graphicData/c:chart', $node);
        if ($chartElement->hasAttribute('r:id') && isset($this->arrayRels[$this->fileRels][$chartElement->getAttribute('r:id')]['Target'])) {
            $pathImage = 'ppt/slides/' . $this->arrayRels[$this->fileRels][$chartElement->getAttribute('r:id')]['Target'];
            $pathImage = explode('/', $pathImage);
            foreach ($pathImage as $key => $partPath) {
                if ('..' == $partPath) {
                    unset($pathImage[$key - 1], $pathImage[$key]);
                }
            }
            $pathChart = implode('/', $pathImage);
            $fileChart = $this->oZip->getFromName($pathChart);
            if (false !== $fileChart) {
                $xmlReader = new XMLReader();
                // @phpstan-ignore-next-line
                if ($xmlReader->getDomFromString($fileChart)) {
                    if ($oElement = $xmlReader->getElement('/c:chartSpace/c:chart/c:autoTitleDeleted')) {
                        $oShape->getTitle()->setVisible(false);
                    }

                    $shapeType = $this->loadTypeChart($xmlReader);
                    if ($shapeType instanceof Chart\Type\AbstractType) {
                        $oShape->getPlotArea()->setType($shapeType);
                    }

                    if ($oElement = $xmlReader->getElement('/c:chartSpace/c:chart/c:plotArea/c:catAx')) {
                        if ($elementOrientation = $xmlReader->getElement('c:scaling/c:orientation', $oElement)) {
                            $oShape->getPlotArea()->getAxisX()->setIsReversedOrder(
                                (bool) ($elementOrientation->getAttribute('val') === 'maxMin')
                            );
                        }
                        if ($elementDelete = $xmlReader->getElement('c:delete', $oElement)) {
                            $oShape->getPlotArea()->getAxisX()->setIsVisible(
                                (bool) ($elementDelete->getAttribute('val') === '0')
                            );
                        }
                        if ($elementMajorTickMark = $xmlReader->getElement('c:majorTickMark', $oElement)) {
                            $oShape->getPlotArea()->getAxisX()->setMajorTickMark($elementMajorTickMark->getAttribute('val'));
                        }
                        if ($elementMinorTickMark = $xmlReader->getElement('c:minorTickMark', $oElement)) {
                            $oShape->getPlotArea()->getAxisX()->setMinorTickMark($elementMinorTickMark->getAttribute('val'));
                        }
                        if ($elementTickLabelPosition = $xmlReader->getElement('c:tickLblPos', $oElement)) {
                            $oShape->getPlotArea()->getAxisX()->setTickLabelPosition($elementTickLabelPosition->getAttribute('val'));
                        }
                        if ($elementCrosses = $xmlReader->getElement('c:crosses', $oElement)) {
                            $oShape->getPlotArea()->getAxisX()->setCrossesAt($elementCrosses->getAttribute('val'));
                        }

                        if ($elementFill = $xmlReader->getElement('c:spPr', $oElement)) {
                            $outline = $this->loadStyleOutline($xmlReader, $elementFill);
                            if ($outline) {
                                $oShape->getPlotArea()->getAxisX()->setOutline($outline);
                            }
                        }

                        $this->loadAxisText($xmlReader, $oElement, $oShape->getPlotArea()->getAxisX());
                    }

                    if ($oElement = $xmlReader->getElement('/c:chartSpace/c:chart/c:plotArea/c:valAx')) {
                        if ($elementOrientation = $xmlReader->getElement('c:scaling/c:orientation', $oElement)) {
                            $oShape->getPlotArea()->getAxisY()->setIsReversedOrder(
                                (bool) ($elementOrientation->getAttribute('val') === 'maxMin')
                            );
                        }
                        if ($elementDelete = $xmlReader->getElement('c:delete', $oElement)) {
                            $oShape->getPlotArea()->getAxisY()->setIsVisible(
                                (bool) ($elementDelete->getAttribute('val') === '0')
                            );
                        }
                        if ($elementMajorTickMark = $xmlReader->getElement('c:majorTickMark', $oElement)) {
                            $oShape->getPlotArea()->getAxisY()->setMajorTickMark($elementMajorTickMark->getAttribute('val'));
                        }
                        if ($elementMinorTickMark = $xmlReader->getElement('c:minorTickMark', $oElement)) {
                            $oShape->getPlotArea()->getAxisY()->setMinorTickMark($elementMinorTickMark->getAttribute('val'));
                        }
                        if ($elementTickLabelPosition = $xmlReader->getElement('c:tickLblPos', $oElement)) {
                            $oShape->getPlotArea()->getAxisY()->setTickLabelPosition($elementTickLabelPosition->getAttribute('val'));
                        }
                        if ($elementCrosses = $xmlReader->getElement('c:crosses', $oElement)) {
                            $oShape->getPlotArea()->getAxisY()->setCrossesAt($elementCrosses->getAttribute('val'));
                        }
                        if ($elementFill = $xmlReader->getElement('c:spPr', $oElement)) {
                            if ($outline = $this->loadStyleOutline($xmlReader, $elementFill)) {
                                $oShape->getPlotArea()->getAxisY()->setOutline($outline);
                            }
                        }

                        $this->loadAxisText($xmlReader, $oElement, $oShape->getPlotArea()->getAxisY());
                    }

                    if ($oElement = $xmlReader->getElement('/c:chartSpace/c:chart/c:legend')) {
                        $oShape->getLegend()->setVisible(true);

                        if ($elementLegendPos = $xmlReader->getElement('c:legendPos', $oElement)) {
                            $oShape->getLegend()->setPosition($elementLegendPos->getAttribute('val'));
                        }
                    } else {
                        $oShape->getLegend()->setVisible(false);
                    }

                    if ($oElement = $xmlReader->getElement('/c:chartSpace/c:chart/c:dispBlanksAs')) {
                        $oShape->setDisplayBlankAs($oElement->getAttribute('val'));
                    }
                }
            }
            $oContainer->addShape($oShape);
        }
    }

    /**
     * Read the plot of a chart, whichever of the nine kinds it is.
     *
     * `c:plotArea` holds exactly one of these elements, and the name of it is the whole of what
     * says which chart this is: a doughnut and a pie carry the same series, the same categories and
     * the same labels, and differ by the element they sit in.
     */
    protected function loadTypeChart(XMLReader $xmlReader): ?Chart\Type\AbstractType
    {
        foreach (self::CHART_TYPES as $name => $class) {
            $oElement = $xmlReader->getElement('/c:chartSpace/c:chart/c:plotArea/' . $name);
            if (!$oElement instanceof DOMElement) {
                continue;
            }

            $shapeType = new $class();
            $this->loadTypeChartProperties($xmlReader, $oElement, $shapeType);

            foreach ($xmlReader->getElements('c:ser', $oElement) as $elementSerie) {
                if ($elementSerie instanceof DOMElement) {
                    $shapeType->addSeries($this->loadSeries($xmlReader, $elementSerie));
                }
            }

            return $shapeType;
        }

        return null;
    }

    /**
     * The settings that belong to the kind of chart this is, rather than to its series.
     *
     * Read by what the type can be told rather than by which element it came from: `c:gapWidth` is
     * written by the bar charts and by the doughnut, and only the first has somewhere to put it.
     */
    protected function loadTypeChartProperties(XMLReader $xmlReader, DOMElement $oElement, Chart\Type\AbstractType $shapeType): void
    {
        if ($shapeType instanceof Chart\Type\AbstractTypeBar) {
            $element = $xmlReader->getElement('c:barDir', $oElement);
            if ($element instanceof DOMElement) {
                $shapeType->setBarDirection($element->getAttribute('val'));
            }

            $element = $xmlReader->getElement('c:grouping', $oElement);
            if ($element instanceof DOMElement) {
                $shapeType->setBarGrouping($element->getAttribute('val'));
            }

            $element = $xmlReader->getElement('c:gapWidth', $oElement);
            if ($element instanceof DOMElement) {
                $shapeType->setGapWidthPercent((int) $element->getAttribute('val'));
            }

            $element = $xmlReader->getElement('c:overlap', $oElement);
            if ($element instanceof DOMElement) {
                $shapeType->setOverlapWidthPercent((int) $element->getAttribute('val'));
            }
        }

        if ($shapeType instanceof Chart\Type\Doughnut) {
            $element = $xmlReader->getElement('c:holeSize', $oElement);
            if ($element instanceof DOMElement) {
                $shapeType->setHoleSize((int) $element->getAttribute('val'));
            }

            $element = $xmlReader->getElement('c:firstSliceAng', $oElement);
            if ($element instanceof DOMElement) {
                $shapeType->setFirstSliceAngle((int) $element->getAttribute('val'));
            }
        }

        // Written inside the first series, and held on the type: a pie explodes as a whole, and a
        // line is smooth or is not, whatever the file repeats for every series it has.
        if ($shapeType instanceof Chart\Type\AbstractTypePie) {
            $element = $xmlReader->getElement('c:ser/c:explosion', $oElement);
            if ($element instanceof DOMElement) {
                $shapeType->setExplosion((int) $element->getAttribute('val'));
            }
        }

        if ($shapeType instanceof Chart\Type\AbstractTypeLine) {
            $element = $xmlReader->getElement('c:ser/c:smooth', $oElement);
            if ($element instanceof DOMElement) {
                $shapeType->setIsSmooth((bool) $element->getAttribute('val'));
            }
        }
    }

    /**
     * The points a series holds, from whichever element of it holds them.
     *
     * A chart written with a worksheet points at the cells and keeps a copy of what it last read
     * there (`c:numRef/c:numCache`); one written without a worksheet holds the points itself
     * (`c:numLit`). The two carry the same `c:ptCount` and the same `c:pt`. Whether they are called
     * `str` or `num` follows the values rather than the axis, so both are looked for on both. A
     * scatter chart calls its two axes `c:xVal` and `c:yVal` where every other chart says `c:cat`
     * and `c:val`, and means the same by them.
     */
    private function loadSeriesPoints(XMLReader $xmlReader, DOMElement $elementSerie, string ...$names): ?DOMElement
    {
        foreach ($names as $name) {
            foreach (['c:strRef/c:strCache', 'c:numRef/c:numCache', 'c:strLit', 'c:numLit'] as $holder) {
                $element = $xmlReader->getElement($name . '/' . $holder, $elementSerie);
                if ($element instanceof DOMElement) {
                    return $element;
                }
            }
        }

        return null;
    }

    /**
     * Read one `c:ser`, which is the same element whatever chart holds it.
     */
    protected function loadSeries(XMLReader $xmlReader, DOMElement $elementSerie): Chart\Series
    {
        $series = new Chart\Series();
        // The name of the series is a cell of the worksheet when there is one, and the text itself
        // when there is not.
        $elementTitle = $xmlReader->getElement('c:tx/c:strRef/c:strCache/c:pt/c:v', $elementSerie)
            ?? $xmlReader->getElement('c:tx/c:v', $elementSerie);
        if ($elementTitle instanceof DOMElement) {
            $series->setTitle($elementTitle->nodeValue);
        }

        $numPoints = 0;
        $elementCategory = $this->loadSeriesPoints($xmlReader, $elementSerie, 'c:cat', 'c:xVal');
        if ($elementCategoryNumPoints = $xmlReader->getElement('c:ptCount', $elementCategory)) {
            $numPoints = (int) $elementCategoryNumPoints->getAttribute('val');
        }
        $elementValue = $this->loadSeriesPoints($xmlReader, $elementSerie, 'c:val', 'c:yVal');
        for ($inc = 0; $inc < $numPoints; ++$inc) {
            $key = '';
            $val = '0';
            if ($subElementCategory = $xmlReader->getElement('c:pt[@idx="' . $inc . '"]/c:v', $elementCategory)) {
                $key = $subElementCategory->nodeValue;
            }
            if ($subElementValue = $xmlReader->getElement('c:pt[@idx="' . $inc . '"]/c:v', $elementValue)) {
                $val = $subElementValue->nodeValue;
            }
            $series->addValue($key, $val);
        }

        if ($elementFill = $xmlReader->getElement('c:spPr', $elementSerie)) {
            $series->setFill(
                $this->loadStyleFill($xmlReader, $elementFill)
            );
        }

        if ($elementFill = $xmlReader->getElement('a:ln', $elementSerie)) {
            $series->setOutline(
                $this->loadStyleOutline($xmlReader, $elementFill)
            );
        }

        $this->loadSeriesDataPoints($xmlReader, $elementSerie, $series);

        // the plate the data labels sit on, which is a `c:spPr` of its own inside `c:dLbls`
        if ($elementFill = $xmlReader->getElement('c:dLbls/c:spPr', $elementSerie)) {
            $series->setLabelFill(
                $this->loadStyleFill($xmlReader, $elementFill)
            );
        }

        if ($elementShowLegendKey = $xmlReader->getElement('c:dLbls/c:showLegendKey', $elementSerie)) {
            $series->setShowLegendKey((bool) $elementShowLegendKey->getAttribute('val'));
        }

        if ($elementShowVal = $xmlReader->getElement('c:dLbls/c:showVal', $elementSerie)) {
            $series->setShowValue((bool) $elementShowVal->getAttribute('val'));
        }

        if ($elementShowCatName = $xmlReader->getElement('c:dLbls/c:showCatName', $elementSerie)) {
            $series->setShowCategoryName((bool) $elementShowCatName->getAttribute('val'));
        }

        if ($elementShowSerName = $xmlReader->getElement('c:dLbls/c:showSerName', $elementSerie)) {
            $series->setShowSeriesName((bool) $elementShowSerName->getAttribute('val'));
        }

        if ($elementShowPercent = $xmlReader->getElement('c:dLbls/c:showPercent', $elementSerie)) {
            $series->setShowPercentage((bool) $elementShowPercent->getAttribute('val'));
        }

        if ($elementShowLeaderLines = $xmlReader->getElement('c:dLbls/c:showLeaderLines', $elementSerie)) {
            $series->setShowLeaderLines((bool) $elementShowLeaderLines->getAttribute('val'));
        }

        return $series;
    }

    /**
     * @param Cell|RichText $oShape
     */
    protected function loadParagraph(XMLReader $document, DOMElement $oElement, $oShape): void
    {
        // Core
        $oParagraph = $oShape->createParagraph();
        $oParagraph->setRichTextElements([]);

        $oSubElement = $document->getElement('a:pPr', $oElement);
        if ($oSubElement instanceof DOMElement) {
            if ($oSubElement->hasAttribute('algn')) {
                $oParagraph->getAlignment()->setHorizontal($oSubElement->getAttribute('algn'));
            }
            if ($oSubElement->hasAttribute('fontAlgn')) {
                $oParagraph->getAlignment()->setVertical($oSubElement->getAttribute('fontAlgn'));
            }
            if ($oSubElement->hasAttribute('marL')) {
                $oParagraph->getAlignment()->setMarginLeft(CommonDrawing::emuToPixels((int) $oSubElement->getAttribute('marL')));
            }
            if ($oSubElement->hasAttribute('marR')) {
                $oParagraph->getAlignment()->setMarginRight(CommonDrawing::emuToPixels((int) $oSubElement->getAttribute('marR')));
            }
            if ($oSubElement->hasAttribute('indent')) {
                $oParagraph->getAlignment()->setIndent((int) CommonDrawing::emuToPixels((int) $oSubElement->getAttribute('indent')));
            }
            if ($oSubElement->hasAttribute('lvl')) {
                $oParagraph->getAlignment()->setLevel((int) $oSubElement->getAttribute('lvl'));
            }
            if ($oSubElement->hasAttribute('rtl')) {
                $oParagraph->getAlignment()->setIsRTL((bool) $oSubElement->getAttribute('rtl'));
            }

            $oElementLineSpacingPoints = $document->getElement('a:lnSpc/a:spcPts', $oSubElement);
            if ($oElementLineSpacingPoints instanceof DOMElement) {
                $oParagraph->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_POINT);
                $oParagraph->setLineSpacing((int) ((int) $oElementLineSpacingPoints->getAttribute('val') / 100));
            }
            $oElementLineSpacingPercent = $document->getElement('a:lnSpc/a:spcPct', $oSubElement);
            if ($oElementLineSpacingPercent instanceof DOMElement) {
                $oParagraph->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_PERCENT);
                $oParagraph->setLineSpacing((int) ((int) $oElementLineSpacingPercent->getAttribute('val') / 1000));
            }
            $oElementSpacingBefore = $document->getElement('a:spcBef/a:spcPts', $oSubElement);
            if ($oElementSpacingBefore instanceof DOMElement) {
                $oParagraph->setSpacingBefore((int) ((int) $oElementSpacingBefore->getAttribute('val') / 100));
            }
            $oElementSpacingAfter = $document->getElement('a:spcAft/a:spcPts', $oSubElement);
            if ($oElementSpacingAfter instanceof DOMElement) {
                $oParagraph->setSpacingAfter((int) ((int) $oElementSpacingAfter->getAttribute('val') / 100));
            }

            $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_NONE);

            $oElementBuFont = $document->getElement('a:buFont', $oSubElement);
            if ($oElementBuFont instanceof DOMElement) {
                if ($oElementBuFont->hasAttribute('typeface')) {
                    $oParagraph->getBulletStyle()->setBulletFont($oElementBuFont->getAttribute('typeface'));
                }
            }
            $oElementBuChar = $document->getElement('a:buChar', $oSubElement);
            if ($oElementBuChar instanceof DOMElement) {
                $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
                if ($oElementBuChar->hasAttribute('char')) {
                    $oParagraph->getBulletStyle()->setBulletChar($oElementBuChar->getAttribute('char'));
                }
            }
            $oElementBuAutoNum = $document->getElement('a:buAutoNum', $oSubElement);
            if ($oElementBuAutoNum instanceof DOMElement) {
                $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_NUMERIC);
                if ($oElementBuAutoNum->hasAttribute('type')) {
                    $oParagraph->getBulletStyle()->setBulletNumericStyle($oElementBuAutoNum->getAttribute('type'));
                }
                if ($oElementBuAutoNum->hasAttribute('startAt') && 1 != $oElementBuAutoNum->getAttribute('startAt')) {
                    $oParagraph->getBulletStyle()->setBulletNumericStartAt($oElementBuAutoNum->getAttribute('startAt'));
                }
            }
            $oElementBuClr = $document->getElement('a:buClr', $oSubElement);
            if ($oElementBuClr instanceof DOMElement) {
                $oColor = new Color();
                /**
                 * @todo Create protected for reading Color
                 */
                $oElementColor = $document->getElement('a:srgbClr', $oElementBuClr);
                if ($oElementColor instanceof DOMElement) {
                    $oColor->setRGB($oElementColor->hasAttribute('val') ? $oElementColor->getAttribute('val') : null);
                }
                $oParagraph->getBulletStyle()->setBulletColor($oColor);
            }
        }
        $arraySubElements = $document->getElements('(a:r|a:br|a:fld)', $oElement);
        foreach ($arraySubElements as $oSubElement) {
            if (!($oSubElement instanceof DOMElement)) {
                continue;
            }
            if ('a:br' == $oSubElement->tagName) {
                $oParagraph->createBreak();
            }
            // A field is a run whose text the application recomputes -- a slide number, a date.
            // `CT_TextField` holds the same `a:rPr` and `a:t` as a run does (ECMA-376), so it is
            // read the same way: what the field says is a stand-in, but how it is styled is not.
            if ('a:r' == $oSubElement->tagName || 'a:fld' == $oSubElement->tagName) {
                $oElementrPr = $document->getElement('a:rPr', $oSubElement);

                // `a:rPr` is optional (ECMA-376, CT_RegularTextRun), and Keynote's export leaves it
                // out, so the run is made before its properties are looked at rather than by them
                $oText = 'a:fld' == $oSubElement->tagName
                    ? $oParagraph->createField($oSubElement->getAttribute('type'))
                    : $oParagraph->createTextRun();
                if ($oElementrPr instanceof DOMElement) {
                    if ($oElementrPr->hasAttribute('b')) {
                        $att = $oElementrPr->getAttribute('b');
                        $oText->getFont()->setBold('true' == $att || '1' == $att ? true : false);
                    }
                    if ($oElementrPr->hasAttribute('i')) {
                        $att = $oElementrPr->getAttribute('i');
                        $oText->getFont()->setItalic('true' == $att || '1' == $att ? true : false);
                    }
                    if ($oElementrPr->hasAttribute('strike')) {
                        $oText->getFont()->setStrikethrough($oElementrPr->getAttribute('strike'));
                    }
                    if ($oElementrPr->hasAttribute('sz')) {
                        $oText->getFont()->setSize((int) ((int) $oElementrPr->getAttribute('sz') / 100));
                    }
                    if ($oElementrPr->hasAttribute('u')) {
                        $oText->getFont()->setUnderline($oElementrPr->getAttribute('u'));
                    }
                    if ($oElementrPr->hasAttribute('cap')) {
                        $oText->getFont()->setCapitalization($oElementrPr->getAttribute('cap'));
                    }
                    if ($oElementrPr->hasAttribute('lang')) {
                        $oText->setLanguage($oElementrPr->getAttribute('lang'));
                    }
                    if ($oElementrPr->hasAttribute('baseline')) {
                        $oText->getFont()->setBaseline((int) $oElementrPr->getAttribute('baseline'));
                    }
                    // Color
                    $oElementSrgbClr = $document->getElement('a:solidFill/a:srgbClr', $oElementrPr);
                    if ($oElementSrgbClr instanceof DOMElement && $oElementSrgbClr->hasAttribute('val')) {
                        $oColor = new Color();
                        $oColor->setRGB($oElementSrgbClr->getAttribute('val'));
                        $oText->getFont()->setColor($oColor);
                    }
                    // Hyperlink
                    $oElementHlinkClick = $document->getElement('a:hlinkClick', $oElementrPr);
                    if ($oElementHlinkClick instanceof DOMElement) {
                        $oText->setHyperlink(
                            $this->loadHyperlink($document, $oElementHlinkClick, $oText->getHyperlink())
                        );
                    }

                    // Font
                    $oElementFontFormat = null;
                    $oElementFontFormatComplexScript = $document->getElement('a:cs', $oElementrPr);
                    if ($oElementFontFormatComplexScript instanceof DOMElement) {
                        $oText->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
                        $oElementFontFormat = $oElementFontFormatComplexScript;
                    }
                    $oElementFontFormatEastAsian = $document->getElement('a:ea', $oElementrPr);
                    if ($oElementFontFormatEastAsian instanceof DOMElement) {
                        $oText->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
                        $oElementFontFormat = $oElementFontFormatEastAsian;
                    }
                    $oElementFontFormatLatin = $document->getElement('a:latin', $oElementrPr);
                    if ($oElementFontFormatLatin instanceof DOMElement) {
                        $oText->getFont()->setFormat(Font::FORMAT_LATIN);
                        $oElementFontFormat = $oElementFontFormatLatin;
                    }
                    if ($oElementFontFormat instanceof DOMElement && $oElementFontFormat->hasAttribute('typeface')) {
                        $oText->getFont()->setName($oElementFontFormat->getAttribute('typeface'));
                    }
                    // Font definition
                    $oElementFont = $document->getElement('a:latin', $oElementrPr);
                    if ($oElementFont instanceof DOMElement) {
                        if ($oElementFont->hasAttribute('typeface')) {
                            $oText->getFont()->setName($oElementFont->getAttribute('typeface'));
                        }
                        if ($oElementFont->hasAttribute('panose')) {
                            $oText->getFont()->setPanose($oElementFont->getAttribute('panose'));
                        }
                        if ($oElementFont->hasAttribute('pitchFamily')) {
                            $oText->getFont()->setPitchFamily((int) $oElementFont->getAttribute('pitchFamily'));
                        }
                        if ($oElementFont->hasAttribute('charset')) {
                            $charset = (int) $oElementFont->getAttribute('charset');
                            $oText->getFont()->setCharset($charset < 0 ? $charset + 256 : $charset);
                        }
                    }
                }
                // The run now exists whether or not it had properties, so the text it holds has
                // to be looked for rather than assumed: a run without `a:t` is malformed, but it
                // is the reader that would raise the error
                $oSubSubElement = $document->getElement('a:t', $oSubElement);
                if ($oSubSubElement instanceof DOMElement) {
                    $oText->setText($oSubSubElement->nodeValue);
                }
            }
        }
    }

    protected function loadHyperlink(XMLReader $xmlReader, DOMElement $element, Hyperlink $hyperlink): Hyperlink
    {
        if ($element->hasAttribute('tooltip')) {
            $hyperlink->setTooltip($element->getAttribute('tooltip'));
        }
        if ($element->hasAttribute('r:id') && isset($this->arrayRels[$this->fileRels][$element->getAttribute('r:id')]['Target'])) {
            $target = $this->arrayRels[$this->fileRels][$element->getAttribute('r:id')]['Target'];
            // A link to another slide is a relationship to its part, and the action says so. Without
            // this the slide number is lost and only the raw part name, `slide2.xml`, comes back.
            if ('ppaction://hlinksldjump' === $element->getAttribute('action')
                && 1 === preg_match('/slide(\d+)\.xml$/', $target, $matches)) {
                $hyperlink->setSlideNumber((int) $matches[1]);
            } else {
                $hyperlink->setUrl($target);
            }
        }
        if ($subElementExt = $xmlReader->getElement('a:extLst/a:ext', $element)) {
            if ($subElementExt->hasAttribute('uri') && $subElementExt->getAttribute('uri') == '{A12FA001-AC4F-418D-AE19-62706E023703}') {
                $hyperlink->setIsTextColorUsed(true);
            }
        }

        return $hyperlink;
    }

    protected function loadStyleBorder(XMLReader $xmlReader, DOMElement $oElement, Border $oBorder): void
    {
        if ($oElement->hasAttribute('w')) {
            $oBorder->setLineWidth(CommonDrawing::emuToPixels((int) $oElement->getAttribute('w')));
        }
        if ($oElement->hasAttribute('cmpd')) {
            $oBorder->setLineStyle($oElement->getAttribute('cmpd'));
        }

        $oElementNoFill = $xmlReader->getElement('a:noFill', $oElement);
        if ($oElementNoFill instanceof DOMElement && Border::LINE_SINGLE == $oBorder->getLineStyle()) {
            $oBorder->setLineStyle(Border::LINE_NONE);
        }

        $oElementColor = $xmlReader->getElement('a:solidFill/a:srgbClr', $oElement);
        if ($oElementColor instanceof DOMElement) {
            $oBorder->setColor($this->loadStyleColor($xmlReader, $oElementColor));
        }

        $oElementDashStyle = $xmlReader->getElement('a:prstDash', $oElement);
        if ($oElementDashStyle instanceof DOMElement && $oElementDashStyle->hasAttribute('val')) {
            $oBorder->setDashStyle($oElementDashStyle->getAttribute('val'));
        }
    }

    protected function loadStyleColor(XMLReader $xmlReader, DOMElement $oElement): Color
    {
        $oColor = new Color();
        $oColor->setRGB($oElement->getAttribute('val'));
        $oElementAlpha = $xmlReader->getElement('a:alpha', $oElement);
        if ($oElementAlpha instanceof DOMElement && $oElementAlpha->hasAttribute('val')) {
            // `a:alpha` counts in thousandths of a percent, which is the percent `setAlpha()` takes
            // and turns into the hex pair in front of the colour. Doing that arithmetic here read
            // every alpha but `FF` back as none at all, and wrote it as a single character, which
            // took the first digit of the colour with it.
            $oColor->setAlpha((int) round((int) $oElementAlpha->getAttribute('val') / 1000));
        }

        return $oColor;
    }

    /**
     * Read the fill and the outline carried by the data points of a serie.
     *
     * @param DOMElement $oElement the `c:ser` element of the serie
     */
    protected function loadSeriesDataPoints(XMLReader $xmlReader, DOMElement $oElement, Chart\Series $series): void
    {
        foreach ($xmlReader->getElements('c:dPt', $oElement) as $oElementDataPoint) {
            if (!$oElementDataPoint instanceof DOMElement) {
                continue;
            }
            $oElementIndex = $xmlReader->getElement('c:idx', $oElementDataPoint);
            $oElementProperties = $xmlReader->getElement('c:spPr', $oElementDataPoint);
            if (!$oElementIndex instanceof DOMElement || !$oElementProperties instanceof DOMElement) {
                continue;
            }
            $index = (int) $oElementIndex->getAttribute('val');

            $oFill = $this->loadStyleFill($xmlReader, $oElementProperties);
            if ($oFill) {
                $series->setDataPointFill($index, $oFill);
            } elseif ($xmlReader->getElement('a:noFill', $oElementProperties) instanceof DOMElement) {
                // A data point stating that it has no fill of its own
                $series->setDataPointFill($index, new Fill());
            }

            $oOutline = $this->loadStyleOutline($xmlReader, $oElementProperties);
            if ($oOutline) {
                $series->setDataPointOutline($index, $oOutline);
            }
        }
    }

    protected function loadStyleFill(XMLReader $xmlReader, DOMElement $oElement): ?Fill
    {
        // Gradient fill
        $oElementFill = $xmlReader->getElement('a:gradFill', $oElement);
        if ($oElementFill instanceof DOMElement) {
            $oFill = new Fill();
            $oFill->setFillType(Fill::FILL_GRADIENT_LINEAR);

            $oElementColor = $xmlReader->getElement('a:gsLst/a:gs[@pos="0"]/a:srgbClr', $oElementFill);
            if ($oElementColor instanceof DOMElement && $oElementColor->hasAttribute('val')) {
                $oFill->setStartColor($this->loadStyleColor($xmlReader, $oElementColor));
            }

            $oElementColor = $xmlReader->getElement('a:gsLst/a:gs[@pos="100000"]/a:srgbClr', $oElementFill);
            if ($oElementColor instanceof DOMElement && $oElementColor->hasAttribute('val')) {
                $oFill->setEndColor($this->loadStyleColor($xmlReader, $oElementColor));
            }

            $oRotation = $xmlReader->getElement('a:lin', $oElementFill);
            if ($oRotation instanceof DOMElement && $oRotation->hasAttribute('ang')) {
                $oFill->setRotation(CommonDrawing::angleToDegrees((int) $oRotation->getAttribute('ang')));
            }

            return $oFill;
        }

        // Pattern fill. `prst` is what names the pattern, and it is optional in the schema, so a
        // `a:pattFill` without one names none -- there is no fill type to read it back as, and it
        // is written by nothing that knows what it is doing. Read as the absence of a fill, the
        // way an element this reader does not know is.
        $oElementFill = $xmlReader->getElement('a:pattFill', $oElement);
        if ($oElementFill instanceof DOMElement
            && in_array($oElementFill->getAttribute('prst'), Fill::PATTERN_TYPES, true)) {
            $oFill = new Fill();
            $oFill->setFillType($oElementFill->getAttribute('prst'));

            $oElementColor = $xmlReader->getElement('a:fgClr/a:srgbClr', $oElementFill);
            if ($oElementColor instanceof DOMElement) {
                $oFill->setStartColor($this->loadStyleColor($xmlReader, $oElementColor));
            }

            $oElementColor = $xmlReader->getElement('a:bgClr/a:srgbClr', $oElementFill);
            if ($oElementColor instanceof DOMElement) {
                $oFill->setEndColor($this->loadStyleColor($xmlReader, $oElementColor));
            }

            return $oFill;
        }

        // Solid fill
        $oElementFill = $xmlReader->getElement('a:solidFill', $oElement);
        if ($oElementFill instanceof DOMElement) {
            $oFill = new Fill();
            $oFill->setFillType(Fill::FILL_SOLID);

            $oElementColor = $xmlReader->getElement('a:srgbClr', $oElementFill);
            if ($oElementColor instanceof DOMElement) {
                $oFill->setStartColor($this->loadStyleColor($xmlReader, $oElementColor));
            }

            return $oFill;
        }

        // No fill. This is a fill, and it is not the same thing as the absence of one below:
        // `a:noFill` is what the file says when something asked to stay transparent.
        $oElementFill = $xmlReader->getElement('a:noFill', $oElement);
        if ($oElementFill instanceof DOMElement) {
            $oFill = new Fill();
            $oFill->setFillType(Fill::FILL_NONE);

            return $oFill;
        }

        return null;
    }

    protected function loadStyleOutline(XMLReader $xmlReader, DOMElement $oElement): ?Outline
    {
        if ($element = $xmlReader->getElement('a:ln', $oElement)) {
            $outline = new Outline();

            $outline->setWidth((int) CommonDrawing::emuToPixels((int) $element->getAttribute('w')));

            $fill = $this->loadStyleFill($xmlReader, $element);
            if ($fill) {
                $outline->setFill($fill);
            }

            return $outline;
        }

        return null;
    }

    /**
     * The `a:defRPr` a chart writes for a font: the same attributes a run carries, on an element
     * that styles a whole paragraph rather than a piece of one.
     */
    protected function loadStyleFont(XMLReader $xmlReader, DOMElement $oElement, Font $oFont): void
    {
        if ($oElement->hasAttribute('b')) {
            $oFont->setBold('true' == $oElement->getAttribute('b') || '1' == $oElement->getAttribute('b'));
        }
        if ($oElement->hasAttribute('i')) {
            $oFont->setItalic('true' == $oElement->getAttribute('i') || '1' == $oElement->getAttribute('i'));
        }
        if ($oElement->hasAttribute('strike')) {
            $oFont->setStrikethrough($oElement->getAttribute('strike'));
        }
        if ($oElement->hasAttribute('u')) {
            $oFont->setUnderline($oElement->getAttribute('u'));
        }
        if ($oElement->hasAttribute('sz')) {
            $oFont->setSize((int) ((int) $oElement->getAttribute('sz') / 100));
        }
        if ($oElement->hasAttribute('baseline')) {
            $oFont->setBaseline((int) $oElement->getAttribute('baseline'));
        }
        $oElementColor = $xmlReader->getElement('a:solidFill/a:srgbClr', $oElement);
        if ($oElementColor instanceof DOMElement) {
            $oFont->setColor($this->loadStyleColor($xmlReader, $oElementColor));
        }
        $oElementLatin = $xmlReader->getElement('a:latin', $oElement);
        if ($oElementLatin instanceof DOMElement && $oElementLatin->hasAttribute('typeface')) {
            $oFont->setName($oElementLatin->getAttribute('typeface'));
        }
    }

    /**
     * The text of an axis: the title it was given, how that title is turned and styled, and the
     * font of the tick labels. All four are written by this library and none was read back.
     *
     * @param DOMElement $oElement the `c:catAx` or `c:valAx` element of the axis
     */
    protected function loadAxisText(XMLReader $xmlReader, DOMElement $oElement, Chart\Axis $oAxis): void
    {
        $oElementTitle = $xmlReader->getElement('c:title/c:tx/c:rich', $oElement);
        if ($oElementTitle instanceof DOMElement) {
            $title = '';
            foreach ($xmlReader->getElements('a:p/a:r/a:t', $oElementTitle) as $oElementText) {
                $title .= $oElementText->nodeValue;
            }
            if ('' !== $title) {
                $oAxis->setTitle($title);
            }

            $oElementBody = $xmlReader->getElement('a:bodyPr', $oElementTitle);
            if ($oElementBody instanceof DOMElement && $oElementBody->hasAttribute('rot')) {
                $oAxis->setTitleRotation((int) CommonDrawing::angleToDegrees((int) $oElementBody->getAttribute('rot')));
            }

            $oElementFont = $xmlReader->getElement('a:p/a:pPr/a:defRPr', $oElementTitle);
            if ($oElementFont instanceof DOMElement) {
                $this->loadStyleFont($xmlReader, $oElementFont, $oAxis->getFont());
            }
        }

        $oElementFont = $xmlReader->getElement('c:txPr/a:p/a:pPr/a:defRPr', $oElement);
        if ($oElementFont instanceof DOMElement) {
            $this->loadStyleFont($xmlReader, $oElementFont, $oAxis->getTickLabelFont());
        }
    }

    protected function loadRels(string $fileRels): void
    {
        $sPart = $this->oZip->getFromName($fileRels);
        if (false !== $sPart) {
            $xmlReader = new XMLReader();
            // @phpstan-ignore-next-line
            if ($xmlReader->getDomFromString($sPart)) {
                foreach ($xmlReader->getElements('*') as $oNode) {
                    if (!($oNode instanceof DOMElement)) {
                        continue;
                    }
                    $this->arrayRels[$fileRels][$oNode->getAttribute('Id')] = [
                        'Target' => $oNode->getAttribute('Target'),
                        'Type' => $oNode->getAttribute('Type'),
                    ];
                }
            }
        }
    }

    /**
     * @param AbstractSlide|Note $oSlide
     * @param DOMNodeList<DOMNode> $oElements
     * @param null|ShapeContainerInterface $oContainer where to put the shapes; the slide itself unless we are inside a group
     *
     * @internal param $baseFile
     */
    protected function loadSlideShapes(XMLReader $document, $oSlide, DOMNodeList $oElements, XMLReader $xmlReader, ?ShapeContainerInterface $oContainer = null): void
    {
        $oContainer = $oContainer ?? $oSlide;
        foreach ($oElements as $oNode) {
            if (!($oNode instanceof DOMElement)) {
                continue;
            }
            switch ($oNode->tagName) {
                case 'p:graphicFrame':
                    if ($oSlide instanceof AbstractSlide) {
                        if ($document->elementExists('a:graphic/a:graphicData/a:tbl', $oNode)) {
                            $this->loadShapeTable($xmlReader, $oNode, $oSlide, $oContainer);
                        }
                        if ($document->elementExists('a:graphic/a:graphicData/c:chart', $oNode)) {
                            $this->loadShapeChart($xmlReader, $oNode, $oSlide, $oContainer);
                        }
                    }

                    break;
                case 'p:pic':
                    if ($this->loadImages && $oSlide instanceof AbstractSlide) {
                        $this->loadShapeDrawing($xmlReader, $oNode, $oSlide, $oContainer);
                    }

                    break;
                case 'p:sp':
                    $this->loadShapeRichText($xmlReader, $oNode, $oSlide, $oContainer);

                    break;
                case 'p:grpSp':
                    $this->loadShapeGroup($xmlReader, $oNode, $oSlide, $xmlReader, $oContainer);

                    break;
                default:
                    //throw new FeatureNotImplementedException();
            }
        }
    }
}
