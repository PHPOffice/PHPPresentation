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
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLReader;
use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Shape\Drawing\Base64;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Slide\Background\Color as BackgroundColor;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Shadow;
use ZipArchive;

/**
 * Serialized format reader.
 */
class ODPresentation implements ReaderInterface
{
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
     * @var array<string, array{alignment: null|Alignment, background: null, shadow: null|Shadow, fill: null|Fill, spacingAfter: null|int, spacingBefore: null|int, lineSpacingMode: null, lineSpacing: null, font: null, listStyle: null}>
     */
    protected $arrayStyles = [];

    /**
     * @var array<string, array<string, null|string>>
     */
    protected $arrayCommonStyles = [];

    /**
     * @var XMLReader
     */
    protected $oXMLReader;

    /**
     * @var int
     */
    protected $levelParagraph = 0;

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
            if (is_array($oZip->statName('META-INF/manifest.xml')) && is_array($oZip->statName('mimetype')) && 'application/vnd.oasis.opendocument.presentation' == $oZip->getFromName('mimetype')) {
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
     *
     * @param string $pFilename
     *
     * @return PhpPresentation
     */
    protected function loadFile($pFilename)
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();

        $this->oZip = new ZipArchive();
        $this->oZip->open($pFilename);

        $this->oXMLReader = new XMLReader();
        if (false !== $this->oXMLReader->getDomFromZip($pFilename, 'meta.xml')) {
            $this->loadDocumentProperties();
        }
        $this->oXMLReader = new XMLReader();
        if (false !== $this->oXMLReader->getDomFromZip($pFilename, 'styles.xml')) {
            $this->loadStylesFile();
        }
        $this->oXMLReader = new XMLReader();
        if (false !== $this->oXMLReader->getDomFromZip($pFilename, 'content.xml')) {
            $this->loadSlides();
            $this->loadPresentationProperties();
        }

        return $this->oPhpPresentation;
    }

    /**
     * Read Document Properties.
     */
    protected function loadDocumentProperties(): void
    {
        $arrayProperties = [
            '/office:document-meta/office:meta/meta:initial-creator' => 'setCreator',
            '/office:document-meta/office:meta/dc:creator' => 'setLastModifiedBy',
            '/office:document-meta/office:meta/dc:title' => 'setTitle',
            '/office:document-meta/office:meta/dc:description' => 'setDescription',
            '/office:document-meta/office:meta/dc:subject' => 'setSubject',
            '/office:document-meta/office:meta/meta:keyword' => 'setKeywords',
            '/office:document-meta/office:meta/meta:creation-date' => 'setCreated',
            '/office:document-meta/office:meta/dc:date' => 'setModified',
        ];
        $properties = $this->oPhpPresentation->getDocumentProperties();
        foreach ($arrayProperties as $path => $property) {
            $oElement = $this->oXMLReader->getElement($path);
            if ($oElement instanceof DOMElement) {
                $value = $oElement->nodeValue;
                if (in_array($property, ['setCreated', 'setModified'])) {
                    $dateTime = DateTime::createFromFormat(DateTime::W3C, $value);
                    if (!$dateTime) {
                        $dateTime = new DateTime();
                    }
                    $value = $dateTime->getTimestamp();
                }
                $properties->{$property}($value);
            }
        }

        foreach ($this->oXMLReader->getElements('/office:document-meta/office:meta/meta:user-defined') as $element) {
            if (!($element instanceof DOMElement)
                || !$element->hasAttribute('meta:name')) {
                continue;
            }
            $propertyName = $element->getAttribute('meta:name');
            $propertyValue = (string) $element->nodeValue;
            $propertyType = $element->getAttribute('meta:value-type');
            switch ($propertyType) {
                case 'boolean':
                    $propertyType = DocumentProperties::PROPERTY_TYPE_BOOLEAN;

                    break;
                case 'float':
                    $propertyType = filter_var($propertyValue, FILTER_VALIDATE_INT) === false
                        ? DocumentProperties::PROPERTY_TYPE_FLOAT
                        : DocumentProperties::PROPERTY_TYPE_INTEGER;

                    break;
                case 'date':
                    $propertyType = DocumentProperties::PROPERTY_TYPE_DATE;

                    break;
                case 'string':
                default:
                    $propertyType = DocumentProperties::PROPERTY_TYPE_STRING;

                    break;
            }
            $properties->setCustomProperty($propertyName, $propertyValue, $propertyType);
        }
    }

    /**
     * Extract all slides.
     */
    protected function loadSlides(): void
    {
        foreach ($this->oXMLReader->getElements('/office:document-content/office:automatic-styles/*') as $oElement) {
            if ($oElement instanceof DOMElement && $oElement->hasAttribute('style:name')) {
                $this->loadStyle($oElement);
            }
        }
        foreach ($this->oXMLReader->getElements('/office:document-content/office:body/office:presentation/draw:page') as $oElement) {
            if ($oElement instanceof DOMElement && 'draw:page' == $oElement->nodeName) {
                $this->loadSlide($oElement);
            }
        }
    }

    protected function loadPresentationProperties(): void
    {
        $element = $this->oXMLReader->getElement('/office:document-content/office:body/office:presentation/presentation:settings');
        if ($element instanceof DOMElement) {
            if ($element->getAttribute('presentation:full-screen') === 'false') {
                $this->oPhpPresentation->getPresentationProperties()->setSlideshowType(PresentationProperties::SLIDESHOW_TYPE_BROWSE);
            }
        }
    }

    /**
     * Extract style.
     */
    protected function loadStyle(DOMElement $nodeStyle): bool
    {
        $keyStyle = $nodeStyle->getAttribute('style:name');

        $nodeDrawingPageProps = $this->oXMLReader->getElement('style:drawing-page-properties', $nodeStyle);
        if ($nodeDrawingPageProps instanceof DOMElement) {
            // Read Background Color
            if ($nodeDrawingPageProps->hasAttribute('draw:fill-color') && 'solid' == $nodeDrawingPageProps->getAttribute('draw:fill')) {
                $oBackground = new BackgroundColor();
                $oColor = new Color();
                $oColor->setRGB(substr($nodeDrawingPageProps->getAttribute('draw:fill-color'), -6));
                $oBackground->setColor($oColor);
            }
            // Read Background Image
            if ('bitmap' == $nodeDrawingPageProps->getAttribute('draw:fill') && $nodeDrawingPageProps->hasAttribute('draw:fill-image-name')) {
                $nameStyle = $nodeDrawingPageProps->getAttribute('draw:fill-image-name');
                if (!empty($this->arrayCommonStyles[$nameStyle]) && 'image' == $this->arrayCommonStyles[$nameStyle]['type'] && !empty($this->arrayCommonStyles[$nameStyle]['path'])) {
                    $tmpBkgImg = tempnam(sys_get_temp_dir(), 'PhpPresentationReaderODPBkg');
                    $contentImg = $this->oZip->getFromName($this->arrayCommonStyles[$nameStyle]['path']);
                    file_put_contents($tmpBkgImg, $contentImg);

                    $oBackground = new Image();
                    $oBackground->setPath($tmpBkgImg);
                }
            }
        }

        $nodeGraphicProps = $this->oXMLReader->getElement('style:graphic-properties', $nodeStyle);
        if ($nodeGraphicProps instanceof DOMElement) {
            // Read Shadow
            if ($nodeGraphicProps->hasAttribute('draw:shadow') && 'visible' == $nodeGraphicProps->getAttribute('draw:shadow')) {
                $oShadow = new Shadow();
                $oShadow->setVisible(true);
                if ($nodeGraphicProps->hasAttribute('draw:shadow-color')) {
                    $oShadow->getColor()->setRGB(substr($nodeGraphicProps->getAttribute('draw:shadow-color'), -6));
                }
                if ($nodeGraphicProps->hasAttribute('draw:shadow-opacity')) {
                    $oShadow->setAlpha(100 - (int) substr($nodeGraphicProps->getAttribute('draw:shadow-opacity'), 0, -1));
                }
                if ($nodeGraphicProps->hasAttribute('draw:shadow-offset-x') && $nodeGraphicProps->hasAttribute('draw:shadow-offset-y')) {
                    $offsetX = (float) substr($nodeGraphicProps->getAttribute('draw:shadow-offset-x'), 0, -2);
                    $offsetY = (float) substr($nodeGraphicProps->getAttribute('draw:shadow-offset-y'), 0, -2);
                    $distance = 0;
                    if (0 != $offsetX) {
                        $distance = ($offsetX < 0 ? $offsetX * -1 : $offsetX);
                    } elseif (0 != $offsetY) {
                        $distance = ($offsetY < 0 ? $offsetY * -1 : $offsetY);
                    }
                    $oShadow->setDirection((int) rad2deg(atan2($offsetY, $offsetX)));
                    $oShadow->setDistance(CommonDrawing::centimetersToPixels($distance));
                }
            }
            // Read Fill
            if ($nodeGraphicProps->hasAttribute('draw:fill')) {
                $value = $nodeGraphicProps->getAttribute('draw:fill');

                switch ($value) {
                    case 'none':
                        $oFill = new Fill();
                        $oFill->setFillType(Fill::FILL_NONE);

                        break;
                    case 'solid':
                        $oFill = new Fill();
                        $oFill->setFillType(Fill::FILL_SOLID);
                        if ($nodeGraphicProps->hasAttribute('draw:fill-color')) {
                            $oColor = new Color();
                            $oColor->setRGB(substr($nodeGraphicProps->getAttribute('draw:fill-color'), 1));
                            $oFill->setStartColor($oColor);
                        }

                        break;
                }
            }
        }

        $nodeTextProperties = $this->oXMLReader->getElement('style:text-properties', $nodeStyle);
        if ($nodeTextProperties instanceof DOMElement) {
            $oFont = new Font();
            if ($nodeTextProperties->hasAttribute('fo:color')) {
                $oFont->getColor()->setRGB(substr($nodeTextProperties->getAttribute('fo:color'), -6));
            }
            if ($nodeTextProperties->hasAttribute('fo:text-transform')) {
                switch ($nodeTextProperties->getAttribute('fo:text-transform')) {
                    case 'none':
                        $oFont->setCapitalization(Font::CAPITALIZATION_NONE);

                        break;
                    case 'lowercase':
                        $oFont->setCapitalization(Font::CAPITALIZATION_SMALL);

                        break;
                    case 'uppercase':
                        $oFont->setCapitalization(Font::CAPITALIZATION_ALL);

                        break;
                }
            }
            // Font Latin
            if ($nodeTextProperties->hasAttribute('fo:font-family')) {
                $oFont
                    ->setName($nodeTextProperties->getAttribute('fo:font-family'))
                    ->setFormat(Font::FORMAT_LATIN);
            }
            if ($nodeTextProperties->hasAttribute('fo:font-weight') && 'bold' == $nodeTextProperties->getAttribute('fo:font-weight')) {
                $oFont
                    ->setBold(true)
                    ->setFormat(Font::FORMAT_LATIN);
            }
            if ($nodeTextProperties->hasAttribute('fo:font-size')) {
                $oFont
                    ->setSize((int) substr($nodeTextProperties->getAttribute('fo:font-size'), 0, -2))
                    ->setFormat(Font::FORMAT_LATIN);
            }
            // Font East Asian
            if ($nodeTextProperties->hasAttribute('style:font-family-asian')) {
                $oFont
                    ->setName($nodeTextProperties->getAttribute('style:font-family-asian'))
                    ->setFormat(Font::FORMAT_EAST_ASIAN);
            }
            if ($nodeTextProperties->hasAttribute('style:font-weight-asian') && 'bold' == $nodeTextProperties->getAttribute('style:font-weight-asian')) {
                $oFont
                    ->setBold(true)
                    ->setFormat(Font::FORMAT_EAST_ASIAN);
            }
            if ($nodeTextProperties->hasAttribute('style:font-size-asian')) {
                $oFont
                    ->setSize((int) substr($nodeTextProperties->getAttribute('style:font-size-asian'), 0, -2))
                    ->setFormat(Font::FORMAT_EAST_ASIAN);
            }
            // Font Complex Script
            if ($nodeTextProperties->hasAttribute('style:font-family-complex')) {
                $oFont
                    ->setName($nodeTextProperties->getAttribute('style:font-family-complex'))
                    ->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
            }
            if ($nodeTextProperties->hasAttribute('style:font-weight-complex') && 'bold' == $nodeTextProperties->getAttribute('style:font-weight-complex')) {
                $oFont
                    ->setBold(true)
                    ->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
            }
            if ($nodeTextProperties->hasAttribute('style:font-size-complex')) {
                $oFont
                    ->setSize((int) substr($nodeTextProperties->getAttribute('style:font-size-complex'), 0, -2))
                    ->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
            }
            if ($nodeTextProperties->hasAttribute('style:script-type')) {
                switch ($nodeTextProperties->getAttribute('style:script-type')) {
                    case 'latin':
                        $oFont->setFormat(Font::FORMAT_LATIN);

                        break;
                    case 'asian':
                        $oFont->setFormat(Font::FORMAT_EAST_ASIAN);

                        break;
                    case 'complex':
                        $oFont->setFormat(Font::FORMAT_COMPLEX_SCRIPT);

                        break;
                }
            }
        }

        $nodeParagraphProps = $this->oXMLReader->getElement('style:paragraph-properties', $nodeStyle);
        if ($nodeParagraphProps instanceof DOMElement) {
            if ($nodeParagraphProps->hasAttribute('fo:line-height')) {
                $lineHeightUnit = $this->getExpressionUnit($nodeParagraphProps->getAttribute('fo:margin-bottom'));
                $lineSpacingMode = $lineHeightUnit == '%' ? Paragraph::LINE_SPACING_MODE_PERCENT : Paragraph::LINE_SPACING_MODE_POINT;
                $lineSpacing = $this->getExpressionValue($nodeParagraphProps->getAttribute('fo:margin-bottom'));
            }
            if ($nodeParagraphProps->hasAttribute('fo:margin-bottom')) {
                $spacingAfter = self::sizeToPoint($nodeParagraphProps->getAttribute('fo:margin-bottom'));
            }
            if ($nodeParagraphProps->hasAttribute('fo:margin-top')) {
                $spacingBefore = self::sizeToPoint($nodeParagraphProps->getAttribute('fo:margin-top'));
            }
            $oAlignment = new Alignment();
            if ($nodeParagraphProps->hasAttribute('fo:text-align')) {
                $oAlignment->setHorizontal($nodeParagraphProps->getAttribute('fo:text-align'));
            }
            if ($nodeParagraphProps->hasAttribute('style:writing-mode')) {
                switch ($nodeParagraphProps->getAttribute('style:writing-mode')) {
                    case 'lr-tb':
                    case 'tb-lr':
                    case 'lr':
                        $oAlignment->setIsRTL(false);

                        break;
                    case 'rl-tb':
                    case 'tb-rl':
                    case 'rl':
                        $oAlignment->setIsRTL(false);

                        break;
                    case 'tb':
                    case 'page':
                    default:
                        break;
                }
            }
        }

        if ('text:list-style' == $nodeStyle->nodeName) {
            $arrayListStyle = [];
            foreach ($this->oXMLReader->getElements('text:list-level-style-bullet', $nodeStyle) as $oNodeListLevel) {
                $oAlignment = new Alignment();
                $oBullet = new Bullet();
                $oBullet->setBulletType(Bullet::TYPE_NONE);
                if ($oNodeListLevel instanceof DOMElement) {
                    if ($oNodeListLevel->hasAttribute('text:level')) {
                        $oAlignment->setLevel((int) $oNodeListLevel->getAttribute('text:level') - 1);
                    }
                    if ($oNodeListLevel->hasAttribute('text:bullet-char')) {
                        $oBullet->setBulletChar($oNodeListLevel->getAttribute('text:bullet-char'));
                        $oBullet->setBulletType(Bullet::TYPE_BULLET);
                    }

                    $oNodeListProperties = $this->oXMLReader->getElement('style:list-level-properties', $oNodeListLevel);
                    if ($oNodeListProperties instanceof DOMElement) {
                        if ($oNodeListProperties->hasAttribute('text:min-label-width')) {
                            $oAlignment->setIndent(CommonDrawing::centimetersToPixels((float) substr($oNodeListProperties->getAttribute('text:min-label-width'), 0, -2)));
                        }
                        if ($oNodeListProperties->hasAttribute('text:space-before')) {
                            $iSpaceBefore = CommonDrawing::centimetersToPixels((float) substr($oNodeListProperties->getAttribute('text:space-before'), 0, -2));
                            $iMarginLeft = $iSpaceBefore + $oAlignment->getIndent();
                            $oAlignment->setMarginLeft($iMarginLeft);
                        }
                    }

                    $oNodeTextProperties = $this->oXMLReader->getElement('style:text-properties', $oNodeListLevel);
                    if ($oNodeTextProperties instanceof DOMElement) {
                        if ($oNodeTextProperties->hasAttribute('fo:font-family')) {
                            $oBullet->setBulletFont($oNodeTextProperties->getAttribute('fo:font-family'));
                        }
                    }
                }

                $arrayListStyle[$oAlignment->getLevel()] = [
                    'alignment' => $oAlignment,
                    'bullet' => $oBullet,
                ];
            }
        }

        $this->arrayStyles[$keyStyle] = [
            'alignment' => $oAlignment ?? null,
            'background' => $oBackground ?? null,
            'fill' => $oFill ?? null,
            'font' => $oFont ?? null,
            'shadow' => $oShadow ?? null,
            'listStyle' => $arrayListStyle ?? null,
            'spacingAfter' => $spacingAfter ?? null,
            'spacingBefore' => $spacingBefore ?? null,
            'lineSpacingMode' => $lineSpacingMode ?? null,
            'lineSpacing' => $lineSpacing ?? null,
        ];

        return true;
    }

    /**
     * Read Slide.
     */
    protected function loadSlide(DOMElement $nodeSlide): bool
    {
        // Core
        $this->oPhpPresentation->createSlide();
        $this->oPhpPresentation->setActiveSlideIndex($this->oPhpPresentation->getSlideCount() - 1);
        if ($nodeSlide->hasAttribute('draw:name')) {
            $this->oPhpPresentation->getActiveSlide()->setName($nodeSlide->getAttribute('draw:name'));
        }
        if ($nodeSlide->hasAttribute('draw:style-name')) {
            $keyStyle = $nodeSlide->getAttribute('draw:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                $this->oPhpPresentation->getActiveSlide()->setBackground($this->arrayStyles[$keyStyle]['background']);
            }
        }
        foreach ($this->oXMLReader->getElements('draw:frame', $nodeSlide) as $oNodeFrame) {
            if ($oNodeFrame instanceof DOMElement) {
                if ($this->loadImages && $this->oXMLReader->getElement('draw:image', $oNodeFrame)) {
                    $this->loadShapeDrawing($oNodeFrame);

                    continue;
                }
                if ($this->oXMLReader->getElement('draw:text-box', $oNodeFrame)) {
                    $this->loadShapeRichText($oNodeFrame);

                    continue;
                }
            }
        }

        return true;
    }

    /**
     * Read Shape Drawing.
     */
    protected function loadShapeDrawing(DOMElement $oNodeFrame): void
    {
        // Core
        $mimetype = '';

        $oNodeImage = $this->oXMLReader->getElement('draw:image', $oNodeFrame);
        if ($oNodeImage instanceof DOMElement) {
            if ($oNodeImage->hasAttribute('loext:mime-type')) {
                $mimetype = $oNodeImage->getAttribute('loext:mime-type');
            }
            if ($oNodeImage->hasAttribute('xlink:href')) {
                $sFilename = $oNodeImage->getAttribute('xlink:href');
                // svm = StarView Metafile
                if ('svm' == pathinfo($sFilename, PATHINFO_EXTENSION)) {
                    return;
                }
                $imageFile = $this->oZip->getFromName($sFilename);
            }
        }

        if (empty($imageFile)) {
            return;
        }

        // Contents of file
        if (empty($mimetype)) {
            $shape = new Gd();
            $shape->setImageResource(imagecreatefromstring($imageFile));
        } else {
            $shape = new Base64();
            $shape->setData('data:' . $mimetype . ';base64,' . base64_encode($imageFile));
        }

        $shape->getShadow()->setVisible(false);
        $shape->setName($oNodeFrame->hasAttribute('draw:name') ? $oNodeFrame->getAttribute('draw:name') : '');
        $shape->setDescription($oNodeFrame->hasAttribute('draw:name') ? $oNodeFrame->getAttribute('draw:name') : '');
        $shape->setResizeProportional(false);
        $shape->setWidth($oNodeFrame->hasAttribute('svg:width') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:width'), 0, -2)) : 0);
        $shape->setHeight($oNodeFrame->hasAttribute('svg:height') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:height'), 0, -2)) : 0);
        $shape->setResizeProportional(true);
        $shape->setOffsetX($oNodeFrame->hasAttribute('svg:x') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:x'), 0, -2)) : 0);
        $shape->setOffsetY($oNodeFrame->hasAttribute('svg:y') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:y'), 0, -2)) : 0);

        if ($oNodeFrame->hasAttribute('draw:style-name')) {
            $keyStyle = $oNodeFrame->getAttribute('draw:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                $shape->setShadow($this->arrayStyles[$keyStyle]['shadow']);
                $shape->setFill($this->arrayStyles[$keyStyle]['fill']);
            }
        }

        $this->oPhpPresentation->getActiveSlide()->addShape($shape);
    }

    /**
     * Read Shape RichText.
     */
    protected function loadShapeRichText(DOMElement $oNodeFrame): void
    {
        // Core
        $oShape = $this->oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->setParagraphs([]);

        $oShape->setWidth($oNodeFrame->hasAttribute('svg:width') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:width'), 0, -2)) : 0);
        $oShape->setHeight($oNodeFrame->hasAttribute('svg:height') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:height'), 0, -2)) : 0);
        $oShape->setOffsetX($oNodeFrame->hasAttribute('svg:x') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:x'), 0, -2)) : 0);
        $oShape->setOffsetY($oNodeFrame->hasAttribute('svg:y') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:y'), 0, -2)) : 0);

        foreach ($this->oXMLReader->getElements('draw:text-box/*', $oNodeFrame) as $oNodeParagraph) {
            $this->levelParagraph = 0;
            if ($oNodeParagraph instanceof DOMElement) {
                if ('text:p' == $oNodeParagraph->nodeName) {
                    $this->readParagraph($oShape, $oNodeParagraph);
                }
                if ('text:list' == $oNodeParagraph->nodeName) {
                    $this->readList($oShape, $oNodeParagraph);
                }
            }
        }

        if (count($oShape->getParagraphs()) > 0) {
            $oShape->setActiveParagraph(0);
        }
    }

    /**
     * Read Paragraph.
     */
    protected function readParagraph(RichText $oShape, DOMElement $oNodeParent): void
    {
        $oParagraph = $oShape->createParagraph();
        if ($oNodeParent->hasAttribute('text:style-name')) {
            $keyStyle = $oNodeParent->getAttribute('text:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                if (!empty($this->arrayStyles[$keyStyle]['spacingAfter'])) {
                    $oParagraph->setSpacingAfter($this->arrayStyles[$keyStyle]['spacingAfter']);
                }
                if (!empty($this->arrayStyles[$keyStyle]['spacingBefore'])) {
                    $oParagraph->setSpacingBefore($this->arrayStyles[$keyStyle]['spacingBefore']);
                }
                if (!empty($this->arrayStyles[$keyStyle]['lineSpacingMode'])) {
                    $oParagraph->setLineSpacingMode($this->arrayStyles[$keyStyle]['lineSpacingMode']);
                }
                if (!empty($this->arrayStyles[$keyStyle]['lineSpacing'])) {
                    $oParagraph->setLineSpacing($this->arrayStyles[$keyStyle]['lineSpacing']);
                }
            }
        }
        $oDomList = $this->oXMLReader->getElements('text:span', $oNodeParent);
        $oDomTextNodes = $this->oXMLReader->getElements('text()', $oNodeParent);
        foreach ($oDomTextNodes as $oDomTextNode) {
            if ('' != trim($oDomTextNode->nodeValue)) {
                $oTextRun = $oParagraph->createTextRun();
                $oTextRun->setText(trim($oDomTextNode->nodeValue));
            }
        }
        foreach ($oDomList as $oNodeRichTextElement) {
            if ($oNodeRichTextElement instanceof DOMElement) {
                $this->readParagraphItem($oParagraph, $oNodeRichTextElement);
            }
        }
    }

    /**
     * Read Paragraph Item.
     */
    protected function readParagraphItem(Paragraph $oParagraph, DOMElement $oNodeParent): void
    {
        if ($this->oXMLReader->elementExists('text:line-break', $oNodeParent)) {
            $oParagraph->createBreak();
        } else {
            $oTextRun = $oParagraph->createTextRun();
            if ($oNodeParent->hasAttribute('text:style-name')) {
                $keyStyle = $oNodeParent->getAttribute('text:style-name');
                if (isset($this->arrayStyles[$keyStyle])) {
                    $oTextRun->setFont($this->arrayStyles[$keyStyle]['font']);
                }
            }
            $oTextRunLink = $this->oXMLReader->getElement('text:a', $oNodeParent);
            if ($oTextRunLink instanceof DOMElement) {
                $oTextRun->setText($oTextRunLink->nodeValue);
                if ($oTextRunLink->hasAttribute('xlink:href')) {
                    $oTextRun->getHyperlink()->setUrl($oTextRunLink->getAttribute('xlink:href'));
                }
            } else {
                $oTextRun->setText($oNodeParent->nodeValue);
            }
        }
    }

    /**
     * Read List.
     */
    protected function readList(RichText $oShape, DOMElement $oNodeParent): void
    {
        foreach ($this->oXMLReader->getElements('text:list-item/*', $oNodeParent) as $oNodeListItem) {
            if ($oNodeListItem instanceof DOMElement) {
                if ('text:p' == $oNodeListItem->nodeName) {
                    $this->readListItem($oShape, $oNodeListItem, $oNodeParent);
                }
                if ('text:list' == $oNodeListItem->nodeName) {
                    ++$this->levelParagraph;
                    $this->readList($oShape, $oNodeListItem);
                    --$this->levelParagraph;
                }
            }
        }
    }

    /**
     * Read List Item.
     */
    protected function readListItem(RichText $oShape, DOMElement $oNodeParent, DOMElement $oNodeParagraph): void
    {
        $oParagraph = $oShape->createParagraph();
        if ($oNodeParagraph->hasAttribute('text:style-name')) {
            $keyStyle = $oNodeParagraph->getAttribute('text:style-name');
            if (isset($this->arrayStyles[$keyStyle]) && !empty($this->arrayStyles[$keyStyle]['listStyle'])) {
                $oParagraph->setAlignment($this->arrayStyles[$keyStyle]['listStyle'][$this->levelParagraph]['alignment']);
                $oParagraph->setBulletStyle($this->arrayStyles[$keyStyle]['listStyle'][$this->levelParagraph]['bullet']);
            }
        }
        foreach ($this->oXMLReader->getElements('text:span', $oNodeParent) as $oNodeRichTextElement) {
            if ($oNodeRichTextElement instanceof DOMElement) {
                $this->readParagraphItem($oParagraph, $oNodeRichTextElement);
            }
        }
    }

    /**
     * Load file 'styles.xml'.
     */
    protected function loadStylesFile(): void
    {
        foreach ($this->oXMLReader->getElements('/office:document-styles/office:styles/*') as $oElement) {
            if ($oElement instanceof DOMElement && 'draw:fill-image' == $oElement->nodeName) {
                $this->arrayCommonStyles[$oElement->getAttribute('draw:name')] = [
                    'type' => 'image',
                    'path' => $oElement->hasAttribute('xlink:href') ? $oElement->getAttribute('xlink:href') : null,
                ];
            }
        }
    }

    private function getExpressionUnit(string $expr): string
    {
        if (substr($expr, -1) == '%') {
            return '%';
        }

        return substr($expr, -2);
    }

    private function getExpressionValue(string $expr): string
    {
        if (substr($expr, -1) == '%') {
            return substr($expr, 0, -1);
        }

        return substr($expr, 0, -2);
    }

    /**
     * Transforms a size in CSS format (eg. 10px, 10px, ...) to points.
     */
    protected static function sizeToPoint(string $value): ?float
    {
        if ($value == '0') {
            return 0;
        }
        $matches = [];
        if (preg_match('/^[+-]?([0-9]+\.?[0-9]*)?(px|em|ex|%|in|cm|mm|pt|pc)$/i', $value, $matches)) {
            $size = (float) $matches[1];
            $unit = $matches[2];

            switch ($unit) {
                case 'pt':
                    return $size;
                case 'px':
                    return CommonDrawing::pixelsToPoints((int) $size);
                case 'cm':
                    return CommonDrawing::centimetersToPoints($size);
                case 'mm':
                    return CommonDrawing::centimetersToPoints($size / 10);
                case 'in':
                    return CommonDrawing::inchesToPoints($size);
                case 'pc':
                    return CommonDrawing::picasToPoints($size);
            }
        }

        return null;
    }
}
