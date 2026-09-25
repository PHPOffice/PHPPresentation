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
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing\Base64;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Field;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\Table\Cell;
use PhpOffice\PhpPresentation\Shape\Table\Row;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Slide\Background\Color as BackgroundColor;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Borders;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Style\Shadow;
use ZipArchive;

/**
 * Serialized format reader.
 */
class ODPresentation implements ReaderInterface
{
    /**
     * The kind of field each OpenDocument field element stands for.
     *
     * `text:time` says only that there is no date in it: which of the timed formats it is travels
     * as a data style this reader does not read, so the plainest of them is what comes back.
     *
     * @var array<string, string>
     */
    protected const FIELD_OOXML = [
        'text:page-number' => Field::TYPE_SLIDENUM,
        'text:page-count' => Field::TYPE_SLIDECOUNT,
        'text:date' => Field::TYPE_DATETIME,
        'text:time' => 'datetime10',
        'text:author-name' => 'author',
        'text:file-name' => 'file',
    ];

    /**
     * The reverse of what the Writer spells out: a style, a type and a width back into the single
     * token OOXML names an underline with. `words` is not here -- ODF says it with a mode.
     *
     * @var array<string, string>
     */
    protected const UNDERLINE_OOXML = [
        'dash|single|' => Font::UNDERLINE_DASH,
        'dash|single|bold' => Font::UNDERLINE_DASHHEAVY,
        'long-dash|single|' => Font::UNDERLINE_DASHLONG,
        'long-dash|single|bold' => Font::UNDERLINE_DASHLONGHEAVY,
        'dot-dash|single|' => Font::UNDERLINE_DOTHASH,
        'dot-dash|single|bold' => Font::UNDERLINE_DOTHASHHEAVY,
        'dot-dot-dash|single|' => Font::UNDERLINE_DOTDOTDASH,
        'dot-dot-dash|single|bold' => Font::UNDERLINE_DOTDOTDASHHEAVY,
        'dotted|single|' => Font::UNDERLINE_DOTTED,
        'dotted|single|bold' => Font::UNDERLINE_DOTTEDHEAVY,
        'solid|double|' => Font::UNDERLINE_DOUBLE,
        'solid|single|bold' => Font::UNDERLINE_HEAVY,
        'solid|single|' => Font::UNDERLINE_SINGLE,
        'wave|single|' => Font::UNDERLINE_WAVY,
        'wave|double|' => Font::UNDERLINE_WAVYDOUBLE,
        'wave|single|bold' => Font::UNDERLINE_WAVYHEAVY,
    ];

    /**
     * The position of a chart legend, by the `chart:legend-position` ODF says it with.
     *
     * @var array<string, string>
     */
    protected const CHART_LEGEND_POSITIONS = [
        'bottom' => Chart\Legend::POSITION_BOTTOM,
        'start' => Chart\Legend::POSITION_LEFT,
        'top' => Chart\Legend::POSITION_TOP,
        'top-end' => Chart\Legend::POSITION_TOPRIGHT,
        'end' => Chart\Legend::POSITION_RIGHT,
    ];

    /**
     * What a chart does with an empty cell, by the `chart:treat-empty-cells` ODF says it with.
     *
     * @var array<string, string>
     */
    protected const CHART_BLANKS = [
        'use-zero' => Chart::BLANKAS_ZERO,
        'leave-gap' => Chart::BLANKAS_GAP,
        'ignore' => Chart::BLANKAS_SPAN,
    ];

    /**
     * The marker of a serie, by the `chart:symbol-name` ODF says it with. The Writer says a dot as
     * a circle too, so a circle is all that comes back of either.
     *
     * @var array<string, string>
     */
    protected const CHART_SYMBOLS = [
        'circle' => Chart\Marker::SYMBOL_CIRCLE,
        'horizontal-bar' => Chart\Marker::SYMBOL_DASH,
        'diamond' => Chart\Marker::SYMBOL_DIAMOND,
        'plus' => Chart\Marker::SYMBOL_PLUS,
        'square' => Chart\Marker::SYMBOL_SQUARE,
        'star' => Chart\Marker::SYMBOL_STAR,
        'arrow-up' => Chart\Marker::SYMBOL_TRIANGLE,
        'x' => Chart\Marker::SYMBOL_X,
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
     * Every key `loadStyle()` puts on a style, so that a node naming a style the file does not
     * define is answered with the same shape rather than a missing offset.
     *
     * @var array<int, string>
     */
    protected const STYLE_KEYS = [
        'alignment', 'background', 'columns', 'columnSpacing', 'columnsRTL', 'fill', 'font',
        'shadow', 'listStyle', 'spacingAfter', 'spacingBefore', 'lineSpacingMode', 'lineSpacing',
        'rowHeight', 'borders', 'border', 'insetBottom', 'insetLeft', 'insetRight',
        'insetTop', 'verticalAlignCenter', 'wrap',
    ];

    /**
     * @var array<string, array{alignment: null|Alignment, background: null|BackgroundColor|Image, columns: null|int, columnSpacing: null|int, columnsRTL: null|bool, fill: null|Fill, font: null|Font, language: null|string, shadow: null|Shadow, listStyle: null|array<int, array{alignment: Alignment, bullet: Bullet}>, spacingAfter: null|float, spacingBefore: null|float, lineSpacingMode: null|string, lineSpacing: null|string, rowHeight: null|int, borders: null|Borders, border: null|Border, insetBottom: null|float, insetLeft: null|float, insetRight: null|float, insetTop: null|float, verticalAlignCenter: null|int, wrap: null|string}>
     */
    protected $arrayStyles = [];

    /**
     * @var array<string, array<string, null|string>>
     */
    protected $arrayCommonStyles = [];

    /**
     * The number of every named slide, so that a link can be resolved back to it.
     *
     * @var array<string, int>
     */
    protected $arraySlideNumbers = [];

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
            '/office:document-meta/office:meta/dc:language' => 'setLanguage',
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
        // A link to another slide addresses it by name and can point forwards, so the names are
        // collected before any slide is read
        $this->arraySlideNumbers = [];
        $slideNumber = 0;
        foreach ($this->oXMLReader->getElements('/office:document-content/office:body/office:presentation/draw:page') as $oElement) {
            if ($oElement instanceof DOMElement && 'draw:page' == $oElement->nodeName) {
                ++$slideNumber;
                if ($oElement->hasAttribute('draw:name')) {
                    $this->arraySlideNumbers[$oElement->getAttribute('draw:name')] = $slideNumber;
                }
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
            // Read Columns
            $nodeColumns = $this->oXMLReader->getElement('style:columns', $nodeGraphicProps);
            if ($nodeColumns instanceof DOMElement) {
                if ($nodeColumns->hasAttribute('fo:column-count')) {
                    $columns = (int) $nodeColumns->getAttribute('fo:column-count');
                }
                if ($nodeColumns->hasAttribute('fo:column-gap')) {
                    $columnSpacing = CommonDrawing::centimetersToPixels(
                        (float) substr($nodeColumns->getAttribute('fo:column-gap'), 0, -2)
                    );
                }
            }
            // Read the order of the columns
            if ($nodeGraphicProps->hasAttribute('style:writing-mode')) {
                switch ($nodeGraphicProps->getAttribute('style:writing-mode')) {
                    case 'lr-tb':
                    case 'tb-lr':
                    case 'lr':
                        $columnsRTL = false;

                        break;
                    case 'rl-tb':
                    case 'tb-rl':
                    case 'rl':
                        $columnsRTL = true;

                        break;
                    case 'tb':
                    case 'page':
                    default:
                        break;
                }
            }
            // Read the border, which a shape says as the stroke of its graphic style. OpenDocument
            // has three strokes and no compound line, so a border comes back single or none; the
            // dash it names carries the pattern, which is finer than the stroke itself.
            if ($nodeGraphicProps->hasAttribute('draw:stroke')) {
                $border = new Border();
                if ('none' === $nodeGraphicProps->getAttribute('draw:stroke')) {
                    $border->setLineStyle(Border::LINE_NONE);
                } else {
                    $border->setLineStyle(Border::LINE_SINGLE)->setDashStyle(Border::DASH_SOLID);
                }
                if ('dash' === $nodeGraphicProps->getAttribute('draw:stroke')) {
                    $dashStyle = substr($nodeGraphicProps->getAttribute('draw:stroke-dash'), strlen('strokeDash_'));
                    $border->setDashStyle(in_array($dashStyle, Border::DASH_STYLES, true) ? $dashStyle : Border::DASH_DASH);
                }
                if ($nodeGraphicProps->hasAttribute('svg:stroke-width')) {
                    // rounded because a width goes into the file as centimetres, and a whole
                    // number of points is not a whole number of them
                    $border->setLineWidth(round(CommonDrawing::centimetersToPoints(
                        (float) substr($nodeGraphicProps->getAttribute('svg:stroke-width'), 0, -2)
                    ), 4));
                }
                if ($nodeGraphicProps->hasAttribute('svg:stroke-color')) {
                    $border->setColor(new Color('FF' . substr($nodeGraphicProps->getAttribute('svg:stroke-color'), 1)));
                }
            }
            // Read the insets, which a text box says as the padding of its frame. The conversion is
            // spelled out rather than run through `centimetersToPixels()`, which rounds an inset to
            // a whole pixel, and kept to the six decimals the Writer writes.
            if ($nodeGraphicProps->hasAttribute('fo:padding-bottom')) {
                $insetBottom = round((float) substr($nodeGraphicProps->getAttribute('fo:padding-bottom'), 0, -2) / 2.54 * CommonDrawing::DPI_96, 6);
            }
            if ($nodeGraphicProps->hasAttribute('fo:padding-left')) {
                $insetLeft = round((float) substr($nodeGraphicProps->getAttribute('fo:padding-left'), 0, -2) / 2.54 * CommonDrawing::DPI_96, 6);
            }
            if ($nodeGraphicProps->hasAttribute('fo:padding-right')) {
                $insetRight = round((float) substr($nodeGraphicProps->getAttribute('fo:padding-right'), 0, -2) / 2.54 * CommonDrawing::DPI_96, 6);
            }
            if ($nodeGraphicProps->hasAttribute('fo:padding-top')) {
                $insetTop = round((float) substr($nodeGraphicProps->getAttribute('fo:padding-top'), 0, -2) / 2.54 * CommonDrawing::DPI_96, 6);
            }
            // Read whether the text is centred between the top and the bottom of the frame
            if ($nodeGraphicProps->hasAttribute('draw:textarea-vertical-align')) {
                $verticalAlignCenter = 'middle' === $nodeGraphicProps->getAttribute('draw:textarea-vertical-align')
                    ? RichText::VALIGN_CENTER
                    : RichText::VALIGN_NOTCENTER;
            }
            // Read whether the text wraps inside the frame
            if ($nodeGraphicProps->hasAttribute('fo:wrap-option')) {
                $wrap = 'no-wrap' === $nodeGraphicProps->getAttribute('fo:wrap-option')
                    ? RichText::WRAP_NONE
                    : RichText::WRAP_SQUARE;
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
            $oFont = $this->loadFont($nodeTextProperties);
            $language = $this->loadLanguage($nodeTextProperties, $oFont->getFormat());
        }

        $nodeParagraphProps = $this->oXMLReader->getElement('style:paragraph-properties', $nodeStyle);
        if ($nodeParagraphProps instanceof DOMElement) {
            if ($nodeParagraphProps->hasAttribute('fo:line-height')) {
                $lineHeightUnit = $this->getExpressionUnit($nodeParagraphProps->getAttribute('fo:line-height'));
                $lineSpacingMode = $lineHeightUnit == '%' ? Paragraph::LINE_SPACING_MODE_PERCENT : Paragraph::LINE_SPACING_MODE_POINT;
                $lineSpacing = $this->getExpressionValue($nodeParagraphProps->getAttribute('fo:line-height'));
            }
            // rounded because a spacing goes into the file as centimetres, and a whole number of
            // points is not a whole number of them
            if ($nodeParagraphProps->hasAttribute('fo:margin-bottom')) {
                $spacingAfter = round((float) self::sizeToPoint($nodeParagraphProps->getAttribute('fo:margin-bottom')), 4);
            }
            if ($nodeParagraphProps->hasAttribute('fo:margin-top')) {
                $spacingBefore = round((float) self::sizeToPoint($nodeParagraphProps->getAttribute('fo:margin-top')), 4);
            }
            $oAlignment = new Alignment();
            if ($nodeParagraphProps->hasAttribute('fo:text-align')) {
                switch ($nodeParagraphProps->getAttribute('fo:text-align')) {
                    case 'right':
                        $oAlignment->setHorizontal(Alignment::HORIZONTAL_RIGHT);

                        break;
                    case 'center':
                        $oAlignment->setHorizontal(Alignment::HORIZONTAL_CENTER);

                        break;
                    case 'justify':
                        $oAlignment->setHorizontal(Alignment::HORIZONTAL_JUSTIFY);

                        break;
                    case 'left':
                    default:
                        $oAlignment->setHorizontal(Alignment::HORIZONTAL_LEFT);

                        break;
                }
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
                        $oAlignment->setIsRTL(true);

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

        // A table row carries its height, and a table cell the borders of its four sides -- the
        // ODPresentation Writer puts the latter on `style:paragraph-properties`, where the ODF
        // shorthand `fo:border` stands for all four and the per-side properties override it.
        $nodeTableRowProps = $this->oXMLReader->getElement('style:table-row-properties', $nodeStyle);
        if ($nodeTableRowProps instanceof DOMElement && $nodeTableRowProps->hasAttribute('style:row-height')) {
            // The ODPresentation Writer measures a row height in points -- `Row::setHeight()` is
            // read as pixels by the PowerPoint2007 Writer, but this is the unit to invert here
            $rowHeight = (int) round(CommonDrawing::centimetersToPoints(
                (float) substr($nodeTableRowProps->getAttribute('style:row-height'), 0, -2)
            ));
        }
        if ($nodeParagraphProps instanceof DOMElement) {
            $borders = $this->loadBorders($nodeParagraphProps);
        }

        $this->arrayStyles[$keyStyle] = [
            'alignment' => $oAlignment ?? null,
            'background' => $oBackground ?? null,
            'columns' => $columns ?? null,
            'columnSpacing' => $columnSpacing ?? null,
            'columnsRTL' => $columnsRTL ?? null,
            'fill' => $oFill ?? null,
            'font' => $oFont ?? null,
            'language' => $language ?? null,
            'shadow' => $oShadow ?? null,
            'listStyle' => $arrayListStyle ?? null,
            'spacingAfter' => $spacingAfter ?? null,
            'spacingBefore' => $spacingBefore ?? null,
            'lineSpacingMode' => $lineSpacingMode ?? null,
            'lineSpacing' => $lineSpacing ?? null,
            'rowHeight' => $rowHeight ?? null,
            'borders' => $borders ?? null,
            'border' => $border ?? null,
            'insetBottom' => $insetBottom ?? null,
            'insetLeft' => $insetLeft ?? null,
            'insetRight' => $insetRight ?? null,
            'insetTop' => $insetTop ?? null,
            'verticalAlignCenter' => $verticalAlignCenter ?? null,
            'wrap' => $wrap ?? null,
        ];

        return true;
    }

    /**
     * Read Slide.
     */
    /**
     * Read the four borders an ODF style names.
     *
     * `fo:border` is the shorthand for all four sides; a per-side property overrides it. Each
     * value is the CSS2 shorthand the ODPresentation Writer produces -- a width in points, a
     * style, and a colour that a border naming none leaves out.
     *
     * @return null|Borders null when the style names no border at all
     */
    protected function loadBorders(DOMElement $nodeParagraphProps): ?Borders
    {
        $sides = [
            'fo:border-top' => 'getTop',
            'fo:border-right' => 'getRight',
            'fo:border-bottom' => 'getBottom',
            'fo:border-left' => 'getLeft',
        ];
        $all = $nodeParagraphProps->hasAttribute('fo:border') ? $nodeParagraphProps->getAttribute('fo:border') : null;
        if (null === $all && !array_filter(array_keys($sides), function (string $attribute) use ($nodeParagraphProps): bool {
            return $nodeParagraphProps->hasAttribute($attribute);
        })) {
            return null;
        }

        $oBorders = new Borders();
        foreach ($sides as $attribute => $getter) {
            $value = $nodeParagraphProps->hasAttribute($attribute) ? $nodeParagraphProps->getAttribute($attribute) : $all;
            if (null !== $value) {
                $this->loadBorder($oBorders->$getter(), $value);
            }
        }

        return $oBorders;
    }

    /**
     * Read one ODF `fo:border` value into a border.
     *
     * The CSS2 border styles are fewer than the OOXML line styles, so what comes back is what
     * CSS2 can say: `double` is the only compound line, and the ten dash patterns arrive as
     * `dashed` or `dotted`.
     */
    protected function loadBorder(Border $oBorder, string $value): void
    {
        $parts = explode(' ', $value);
        if (isset($parts[0]) && 'pt' === substr($parts[0], -2)) {
            $oBorder->setLineWidth((float) substr($parts[0], 0, -2) * 1.75);
        }
        switch ($parts[1] ?? '') {
            case 'none':
                $oBorder->setLineStyle(Border::LINE_NONE);

                break;
            case 'double':
                $oBorder->setLineStyle(Border::LINE_DOUBLE)->setDashStyle(Border::DASH_SOLID);

                break;
            case 'dotted':
                $oBorder->setLineStyle(Border::LINE_SINGLE)->setDashStyle(Border::DASH_DOT);

                break;
            case 'dashed':
                $oBorder->setLineStyle(Border::LINE_SINGLE)->setDashStyle(Border::DASH_DASH);

                break;
            default:
                $oBorder->setLineStyle(Border::LINE_SINGLE)->setDashStyle(Border::DASH_SOLID);

                break;
        }
        // A border written without the optional third part named no colour, and takes the one its
        // parent style gives it
        if (isset($parts[2]) && 0 === strpos($parts[2], '#')) {
            $oBorder->setColor(new Color('FF' . substr($parts[2], 1)));
        }
    }

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
                // A table is looked for first: a producer may put a replacement image of the
                // table in the same frame, and the table is what the frame holds
                if ($this->oXMLReader->getElement('table:table', $oNodeFrame)) {
                    $this->loadShapeTable($oNodeFrame);

                    continue;
                }
                // So is a chart, next to which LibreOffice puts the picture of it it last drew
                if ($this->oXMLReader->getElement('draw:object', $oNodeFrame) && $this->loadShapeChart($oNodeFrame)) {
                    continue;
                }
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
        // a line is a shape of the page in its own right, not the content of a frame
        foreach ($this->oXMLReader->getElements('draw:line', $nodeSlide) as $oNodeLine) {
            if ($oNodeLine instanceof DOMElement) {
                $this->loadShapeLine($oNodeLine);
            }
        }
        $this->loadSlideNote($nodeSlide);

        return true;
    }

    /**
     * Read the speaker notes of a slide, the text a producer puts in presentation:notes.
     */
    protected function loadSlideNote(DOMElement $nodeSlide): void
    {
        $note = $this->oPhpPresentation->getActiveSlide()->getNote();

        foreach ($this->oXMLReader->getElements('presentation:notes/draw:frame', $nodeSlide) as $oNodeFrame) {
            if ($oNodeFrame instanceof DOMElement) {
                // An empty text box is the placeholder every slide carries whether it has notes
                // or not, so only a box with something in it becomes a shape
                $oNodeTextBox = $this->oXMLReader->getElement('draw:text-box', $oNodeFrame);
                if ($oNodeTextBox instanceof DOMElement && $oNodeTextBox->hasChildNodes()) {
                    $this->loadShapeRichText($oNodeFrame, $note);
                }
            }
        }
    }

    /**
     * Read the font a `style:text-properties` describes, out of its attributes alone -- which is
     * what lets a chart, whose styles live in a document of their own, read its fonts the same way.
     */
    protected function loadFont(DOMElement $nodeTextProperties): Font
    {
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
        // Italic, spelled once per script the way the family, the size and the weight are
        if ('italic' == $nodeTextProperties->getAttribute('fo:font-style')) {
            $oFont
                ->setItalic(true)
                ->setFormat(Font::FORMAT_LATIN);
        }
        if ('italic' == $nodeTextProperties->getAttribute('style:font-style-asian')) {
            $oFont
                ->setItalic(true)
                ->setFormat(Font::FORMAT_EAST_ASIAN);
        }
        if ('italic' == $nodeTextProperties->getAttribute('style:font-style-complex')) {
            $oFont
                ->setItalic(true)
                ->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        }
        // Underline and strikethrough, one family each for the whole run
        $underlineStyle = $nodeTextProperties->getAttribute('style:text-underline-style');
        if ('' !== $underlineStyle && 'none' !== $underlineStyle) {
            if ('skip-white-space' == $nodeTextProperties->getAttribute('style:text-underline-mode')) {
                $oFont->setUnderline(Font::UNDERLINE_WORDS);
            } else {
                $underlineType = $nodeTextProperties->getAttribute('style:text-underline-type');
                $underlineWidth = $nodeTextProperties->getAttribute('style:text-underline-width');
                $key = $underlineStyle
                    . '|' . ('' === $underlineType ? 'single' : $underlineType)
                    . '|' . ('bold' === $underlineWidth ? 'bold' : '');
                // a file written elsewhere can name a pair OOXML has no word for
                $oFont->setUnderline(self::UNDERLINE_OOXML[$key] ?? Font::UNDERLINE_SINGLE);
            }
        }
        $lineThroughStyle = $nodeTextProperties->getAttribute('style:text-line-through-style');
        if ('' !== $lineThroughStyle && 'none' !== $lineThroughStyle) {
            $oFont->setStrikethrough(
                'double' == $nodeTextProperties->getAttribute('style:text-line-through-type')
                    ? Font::STRIKE_DOUBLE
                    : Font::STRIKE_SINGLE
            );
        }
        $textPosition = $nodeTextProperties->getAttribute('style:text-position');
        if ('' !== $textPosition) {
            $oFont->setBaseline($this->baselineFromTextPosition($textPosition));
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

        return $oFont;
    }

    /**
     * Put back together the language a text style splits: the whole tag where
     * `style:rfc-language-tag` holds it, else the language and the country. The family the style
     * is spelled for is asked first, and the other two after it.
     */
    protected function loadLanguage(DOMElement $nodeTextProperties, string $format): ?string
    {
        $families = [Font::FORMAT_LATIN => '', Font::FORMAT_EAST_ASIAN => '-asian', Font::FORMAT_COMPLEX_SCRIPT => '-complex'];
        $families = [$format => $families[$format] ?? ''] + $families;
        foreach ($families as $suffix) {
            $tag = $nodeTextProperties->getAttribute('style:rfc-language-tag' . $suffix);
            if ('' !== $tag) {
                return $tag;
            }
            $prefix = '' === $suffix ? 'fo:' : 'style:';
            $language = $nodeTextProperties->getAttribute($prefix . 'language' . $suffix);
            if ('' !== $language && 'none' !== $language) {
                $country = $nodeTextProperties->getAttribute($prefix . 'country' . $suffix);

                return $language . ('' !== $country && 'none' !== $country ? '-' . $country : '');
            }
        }

        return null;
    }

    /**
     * Read the description of a shape, the alternative text exposed to assistive
     * technologies. Falls back to the shape name, as written by older versions.
     */
    protected function loadShapeDescription(DOMElement $oNodeFrame): string
    {
        $oNodeDesc = $this->oXMLReader->getElement('svg:desc', $oNodeFrame);
        if ($oNodeDesc instanceof DOMElement) {
            return $oNodeDesc->nodeValue ?? '';
        }

        return $oNodeFrame->hasAttribute('draw:name') ? $oNodeFrame->getAttribute('draw:name') : '';
    }

    /**
     * Read the decorative flag of a shape.
     *
     * @return bool false when the shape says nothing about it
     */
    protected function loadShapeDecorative(DOMElement $oNodeFrame): bool
    {
        if (!$oNodeFrame->hasAttribute('loext:decorative')) {
            return false;
        }

        return 'true' === $oNodeFrame->getAttribute('loext:decorative');
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
        $shape->setDescription($this->loadShapeDescription($oNodeFrame));
        $shape->setDecorative($this->loadShapeDecorative($oNodeFrame));
        $shape->setResizeProportional(false);
        $shape->setWidth($oNodeFrame->hasAttribute('svg:width') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:width'), 0, -2)) : 0);
        $shape->setHeight($oNodeFrame->hasAttribute('svg:height') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:height'), 0, -2)) : 0);
        $shape->setResizeProportional(true);
        $this->loadShapeOffset($shape, $oNodeFrame);

        if ($oNodeFrame->hasAttribute('draw:style-name')) {
            $keyStyle = $oNodeFrame->getAttribute('draw:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                $shape->setShadow($this->arrayStyles[$keyStyle]['shadow']);
                $shape->setFill($this->arrayStyles[$keyStyle]['fill']);
                $this->applyShapeBorder($shape, $this->arrayStyles[$keyStyle]['border']);
            }
        }

        $this->oPhpPresentation->getActiveSlide()->addShape($shape);
    }

    /**
     * Put the border a graphic style names on a shape, where the style named one.
     */
    protected function applyShapeBorder(AbstractShape $shape, ?Border $border): void
    {
        if (null === $border) {
            return;
        }

        $shape->setBorder($border);
    }

    /**
     * Read a Line.
     *
     * ODF says a line by the two points it runs between, so there is nothing to flip: a line that
     * runs right to left simply has an `svg:x2` smaller than its `svg:x1`, which is the negative
     * width the model carries.
     */
    protected function loadShapeLine(DOMElement $oNodeLine): void
    {
        $point = function (string $attribute) use ($oNodeLine): int {
            return $oNodeLine->hasAttribute($attribute)
                ? (int) CommonDrawing::centimetersToPixels((float) substr($oNodeLine->getAttribute($attribute), 0, -2))
                : 0;
        };

        $shape = new Line($point('svg:x1'), $point('svg:y1'), $point('svg:x2'), $point('svg:y2'));
        $shape->setDescription($this->loadShapeDescription($oNodeLine));
        $shape->setDecorative($this->loadShapeDecorative($oNodeLine));

        $this->oPhpPresentation->getActiveSlide()->addShape($shape);
    }

    /**
     * Read a chart, the frame of which embeds an object holding its own `content.xml`: the chart,
     * its styles and the table its series take their values from. A frame embedding anything else
     * is not a chart, and is left out.
     */
    protected function loadShapeChart(DOMElement $oNodeFrame): bool
    {
        $oNodeObject = $this->oXMLReader->getElement('draw:object', $oNodeFrame);
        if (!$oNodeObject instanceof DOMElement) {
            return false;
        }
        $path = preg_replace('#^\./#', '', $oNodeObject->getAttribute('xlink:href')) . '/content.xml';
        $content = $this->oZip->getFromName($path);
        $xmlReader = new XMLReader();
        if (false === $content || false === $xmlReader->getDomFromString($content)) {
            return false;
        }
        $nodeChart = $xmlReader->getElement('/office:document-content/office:body/office:chart/chart:chart');
        if (!$nodeChart instanceof DOMElement) {
            return false;
        }
        $nodePlotArea = $xmlReader->getElement('chart:plot-area', $nodeChart);
        $chartType = $this->loadChartType($xmlReader, $nodeChart, $nodePlotArea);
        if (null === $chartType) {
            return false;
        }

        $shape = new Chart();
        $shape->setName($oNodeFrame->getAttribute('draw:name'));
        $shape->setDescription($this->loadShapeDescription($oNodeFrame));
        $shape->setDecorative($this->loadShapeDecorative($oNodeFrame));
        $shape->setResizeProportional(false);
        $shape->setWidth(CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:width'), 0, -2)));
        $shape->setHeight(CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:height'), 0, -2)));
        $shape->setResizeProportional(true);
        $this->loadShapeOffset($shape, $oNodeFrame);
        $shape->getPlotArea()->setType($chartType);

        $graphicProps = $this->getChartStyle($xmlReader, $nodeChart, 'graphic');
        if ($graphicProps instanceof DOMElement && 'solid' === $graphicProps->getAttribute('draw:fill')) {
            $shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($this->loadChartColor($graphicProps, 'draw:fill-color'));
        }

        // A chart says its title and its legend by holding them: one that has none has none shown
        $nodeTitle = $xmlReader->getElement('chart:title', $nodeChart);
        $shape->getTitle()->setVisible($nodeTitle instanceof DOMElement);
        if ($nodeTitle instanceof DOMElement) {
            $shape->getTitle()->setText($this->loadChartText($xmlReader, $nodeTitle));
            $shape->getTitle()->setOffsetX(CommonDrawing::centimetersToPixels((float) substr($nodeTitle->getAttribute('svg:x'), 0, -2)));
            $shape->getTitle()->setOffsetY(CommonDrawing::centimetersToPixels((float) substr($nodeTitle->getAttribute('svg:y'), 0, -2)));
            $shape->getTitle()->setFont($this->loadChartFont($xmlReader, $nodeTitle));
        }
        $nodeLegend = $xmlReader->getElement('chart:legend', $nodeChart);
        $shape->getLegend()->setVisible($nodeLegend instanceof DOMElement);
        if ($nodeLegend instanceof DOMElement) {
            $shape->getLegend()->setPosition(self::CHART_LEGEND_POSITIONS[$nodeLegend->getAttribute('chart:legend-position')] ?? Chart\Legend::POSITION_RIGHT);
            $shape->getLegend()->setOffsetX(CommonDrawing::centimetersToPixels((float) substr($nodeLegend->getAttribute('svg:x'), 0, -2)));
            $shape->getLegend()->setOffsetY(CommonDrawing::centimetersToPixels((float) substr($nodeLegend->getAttribute('svg:y'), 0, -2)));
            $shape->getLegend()->setFont($this->loadChartFont($xmlReader, $nodeLegend));
        }

        if ($nodePlotArea instanceof DOMElement) {
            $chartProps = $this->getChartStyle($xmlReader, $nodePlotArea, 'chart');
            if ($chartProps instanceof DOMElement && isset(self::CHART_BLANKS[$chartProps->getAttribute('chart:treat-empty-cells')])) {
                $shape->setDisplayBlankAs(self::CHART_BLANKS[$chartProps->getAttribute('chart:treat-empty-cells')]);
            }
            $this->loadChartAxis($xmlReader, $nodePlotArea, 'x', $shape->getPlotArea()->getAxisX());
            $this->loadChartAxis($xmlReader, $nodePlotArea, 'y', $shape->getPlotArea()->getAxisY());
            $this->loadChartSeries($xmlReader, $nodeChart, $nodePlotArea, $chartType);
        }

        $this->oPhpPresentation->getActiveSlide()->addShape($shape);

        return true;
    }

    /**
     * The kind of chart, which `chart:class` names and the plot area says whether it is drawn in
     * three dimensions. A class none of the types draws -- a stock chart, a bubble chart -- is none.
     */
    protected function loadChartType(XMLReader $xmlReader, DOMElement $nodeChart, ?DOMElement $nodePlotArea): ?Chart\Type\AbstractType
    {
        $chartProps = null === $nodePlotArea ? null : $this->getChartStyle($xmlReader, $nodePlotArea, 'chart');
        $is3D = null !== $chartProps && 'true' === $chartProps->getAttribute('chart:three-dimensional');

        switch ($nodeChart->getAttribute('chart:class')) {
            case 'chart:area':
                return new Chart\Type\Area();
            case 'chart:bar':
                $chartType = $is3D ? new Chart\Type\Bar3D() : new Chart\Type\Bar();
                if (null !== $chartProps) {
                    $chartType->setBarDirection('true' === $chartProps->getAttribute('chart:vertical') ? Chart\Type\AbstractTypeBar::DIRECTION_HORIZONTAL : Chart\Type\AbstractTypeBar::DIRECTION_VERTICAL);
                    if ('true' === $chartProps->getAttribute('chart:percentage')) {
                        $chartType->setBarGrouping(Chart\Type\AbstractTypeBar::GROUPING_PERCENTSTACKED);
                    } elseif ('true' === $chartProps->getAttribute('chart:stacked')) {
                        $chartType->setBarGrouping(Chart\Type\AbstractTypeBar::GROUPING_STACKED);
                    }
                }

                return $chartType;
            case 'chart:circle':
                return $is3D ? new Chart\Type\Pie3D() : new Chart\Type\Pie();
            case 'chart:ring':
                return new Chart\Type\Doughnut();
            case 'chart:line':
                $chartType = new Chart\Type\Line();

                break;
            case 'chart:scatter':
                $chartType = new Chart\Type\Scatter();

                break;
            case 'chart:radar':
            case 'chart:filled-radar':
                return new Chart\Type\Radar();
            default:
                return null;
        }
        $chartType->setIsSmooth(null !== $chartProps && in_array($chartProps->getAttribute('chart:interpolation'), ['cubic-spline', 'b-spline'], true));

        return $chartType;
    }

    /**
     * Read an axis: its title, its bounds, its units, where its labels go, its line, its gridlines,
     * and the fonts of its labels and of its title.
     */
    protected function loadChartAxis(XMLReader $xmlReader, DOMElement $nodePlotArea, string $dimension, Chart\Axis $axis): void
    {
        $nodeAxis = $xmlReader->getElement('chart:axis[@chart:dimension="' . $dimension . '"]', $nodePlotArea);
        if (!$nodeAxis instanceof DOMElement) {
            return;
        }

        $nodeTitle = $xmlReader->getElement('chart:title', $nodeAxis);
        if ($nodeTitle instanceof DOMElement) {
            $axis->setTitle($this->loadChartText($xmlReader, $nodeTitle));
            $axis->setFont($this->loadChartFont($xmlReader, $nodeTitle));
            $titleProps = $this->getChartStyle($xmlReader, $nodeTitle, 'chart');
            if ($titleProps instanceof DOMElement && $titleProps->hasAttribute('style:rotation-angle')) {
                $axis->setTitleRotation(-(int) $titleProps->getAttribute('style:rotation-angle'));
            }
        }
        $axis->setTickLabelFont($this->loadChartFont($xmlReader, $nodeAxis));

        $chartProps = $this->getChartStyle($xmlReader, $nodeAxis, 'chart');
        if ($chartProps instanceof DOMElement) {
            if ($chartProps->hasAttribute('chart:minimum')) {
                $axis->setMinBounds((int) $chartProps->getAttribute('chart:minimum'));
            }
            if ($chartProps->hasAttribute('chart:maximum')) {
                $axis->setMaxBounds((int) $chartProps->getAttribute('chart:maximum'));
            }
            if ($chartProps->hasAttribute('chart:interval-major')) {
                $axis->setMajorUnit((float) $chartProps->getAttribute('chart:interval-major'));
                // ODF counts the minor intervals a major one is split in, OOXML the size of one
                $divisor = (int) $chartProps->getAttribute('chart:interval-minor-divisor');
                if ($divisor > 0) {
                    $axis->setMinorUnit((float) $chartProps->getAttribute('chart:interval-major') / $divisor);
                }
            }
            $labelPositions = [
                'near-axis' => Chart\Axis::TICK_LABEL_POSITION_NEXT_TO,
                'outside-end' => Chart\Axis::TICK_LABEL_POSITION_HIGH,
                'outside-start' => Chart\Axis::TICK_LABEL_POSITION_LOW,
            ];
            if (isset($labelPositions[$chartProps->getAttribute('chart:axis-label-position')])) {
                $axis->setTickLabelPosition($labelPositions[$chartProps->getAttribute('chart:axis-label-position')]);
            }
        }

        $graphicProps = $this->getChartStyle($xmlReader, $nodeAxis, 'graphic');
        if ($graphicProps instanceof DOMElement) {
            $axis->setOutline($this->loadChartOutline($graphicProps, true));
        }

        foreach ($xmlReader->getElements('chart:grid', $nodeAxis) as $nodeGrid) {
            if (!$nodeGrid instanceof DOMElement) {
                continue;
            }
            $gridlines = new Chart\Gridlines();
            $graphicProps = $this->getChartStyle($xmlReader, $nodeGrid, 'graphic');
            if ($graphicProps instanceof DOMElement) {
                $gridlines->setOutline($this->loadChartOutline($graphicProps, true));
            }
            if ('minor' === $nodeGrid->getAttribute('chart:class')) {
                $axis->setMinorGridlines($gridlines);
            } else {
                $axis->setMajorGridlines($gridlines);
            }
        }
    }

    /**
     * Read the series, which take their values, their title and their categories out of the table
     * the chart holds, by the cells `chart:series` and `chart:categories` point at. A scatter chart
     * LibreOffice writes points at its X values with `chart:domain` instead.
     */
    protected function loadChartSeries(XMLReader $xmlReader, DOMElement $nodeChart, DOMElement $nodePlotArea, Chart\Type\AbstractType $chartType): void
    {
        $table = $this->loadChartTable($xmlReader, $nodeChart);
        $categories = $this->loadChartRange($table, $xmlReader->getAttribute('table:cell-range-address', $nodePlotArea, 'chart:axis[@chart:dimension="x"]/chart:categories') ?? '');

        foreach ($xmlReader->getElements('chart:series', $nodePlotArea) as $nodeSeries) {
            if (!$nodeSeries instanceof DOMElement) {
                continue;
            }
            $keys = $this->loadChartRange($table, $xmlReader->getAttribute('table:cell-range-address', $nodeSeries, 'chart:domain') ?? '') ?: $categories;
            $values = [];
            foreach ($this->loadChartRange($table, $nodeSeries->getAttribute('chart:values-cell-range-address')) as $index => $value) {
                $values[(string) ($keys[$index] ?? $index + 1)] = $value;
            }
            $title = $this->loadChartRange($table, $nodeSeries->getAttribute('chart:label-cell-address'));
            $series = new Chart\Series((string) ($title[0] ?? ''), $values);

            $chartProps = $this->getChartStyle($xmlReader, $nodeSeries, 'chart');
            if ($chartProps instanceof DOMElement) {
                $this->loadChartSeriesProperties($chartProps, $series, $chartType);
            }
            $graphicProps = $this->getChartStyle($xmlReader, $nodeSeries, 'graphic');
            if ($graphicProps instanceof DOMElement) {
                $fill = $this->loadChartFill($graphicProps);
                if (null !== $fill) {
                    $series->setFill($fill);
                }
                if ($graphicProps->hasAttribute('svg:stroke-color') && !$chartType instanceof Chart\Type\AbstractTypePie) {
                    $series->setOutline($this->loadChartOutline($graphicProps, false));
                }
            }
            $series->setFont($this->loadChartFont($xmlReader, $nodeSeries));

            // A data point that differs from its serie has a style of its own; a run of those that
            // do not is one element saying how many they are
            $index = 0;
            foreach ($xmlReader->getElements('chart:data-point', $nodeSeries) as $nodeDataPoint) {
                if (!$nodeDataPoint instanceof DOMElement) {
                    continue;
                }
                $graphicProps = $this->getChartStyle($xmlReader, $nodeDataPoint, 'graphic');
                if ($graphicProps instanceof DOMElement) {
                    $fill = $this->loadChartFill($graphicProps);
                    if (null !== $fill) {
                        $series->setDataPointFill($index, $fill);
                    }
                    if ($graphicProps->hasAttribute('draw:stroke')) {
                        $series->setDataPointOutline($index, $this->loadChartOutline($graphicProps, false));
                    }
                }
                $index += max(1, (int) $nodeDataPoint->getAttribute('chart:repeated'));
            }

            $chartType->addSeries($series);
        }
    }

    /**
     * Read what the style of a serie says of its labels, its marker and -- for a pie, which explodes
     * as a whole -- its explosion.
     */
    protected function loadChartSeriesProperties(DOMElement $chartProps, Chart\Series $series, Chart\Type\AbstractType $chartType): void
    {
        $labelNumber = $chartProps->getAttribute('chart:data-label-number');
        $series->setShowValue(in_array($labelNumber, ['value', 'value-and-percentage'], true));
        $series->setShowPercentage(in_array($labelNumber, ['percentage', 'value-and-percentage'], true));
        $series->setShowCategoryName('true' === $chartProps->getAttribute('chart:data-label-text'));
        $nodeSeparator = $chartProps->getElementsByTagNameNS('urn:oasis:names:tc:opendocument:xmlns:chart:1.0', 'label-separator')->item(0);
        if ($nodeSeparator instanceof DOMElement) {
            $series->setSeparator($nodeSeparator->getElementsByTagNameNS('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'line-break')->length > 0 ? PHP_EOL : $nodeSeparator->textContent);
        }

        if ($chartType instanceof Chart\Type\AbstractTypePie && $chartProps->hasAttribute('chart:pie-offset') && [] === $chartType->getSeries()) {
            $chartType->setExplosion((int) $chartProps->getAttribute('chart:pie-offset'));
        }

        if ('none' === $chartProps->getAttribute('chart:symbol-type')) {
            $series->getMarker()->setSymbol(Chart\Marker::SYMBOL_NONE);
        } elseif ('named-symbol' === $chartProps->getAttribute('chart:symbol-type')) {
            $symbol = self::CHART_SYMBOLS[$chartProps->getAttribute('chart:symbol-name')] ?? null;
            if (null !== $symbol) {
                $series->getMarker()->setSymbol($symbol);
            }
            if ($chartProps->hasAttribute('chart:symbol-width')) {
                $series->getMarker()->setSize((int) round(CommonDrawing::centimetersToPoints((float) substr($chartProps->getAttribute('chart:symbol-width'), 0, -2))));
            }
        }
    }

    /**
     * The cells of the table a chart holds, by row and column from 1, the header row included: a
     * number as the text of it, which is how a serie holds its values and how the PowerPoint2007
     * Reader reads them -- `NaN` being the empty cell the Writer writes -- and anything else as
     * its text.
     *
     * @return array<int, array<int, null|string>>
     */
    protected function loadChartTable(XMLReader $xmlReader, DOMElement $nodeChart): array
    {
        $table = [];
        $row = 0;
        foreach ($xmlReader->getElements('table:table//table:table-row', $nodeChart) as $nodeRow) {
            if (!$nodeRow instanceof DOMElement) {
                continue;
            }
            ++$row;
            $column = 0;
            foreach ($xmlReader->getElements('table:table-cell', $nodeRow) as $nodeCell) {
                if (!$nodeCell instanceof DOMElement) {
                    continue;
                }
                $value = $this->loadChartText($xmlReader, $nodeCell);
                if ('float' === $nodeCell->getAttribute('office:value-type') || 'percentage' === $nodeCell->getAttribute('office:value-type')) {
                    $value = $nodeCell->getAttribute('office:value');
                    $value = is_numeric($value) ? $value : null;
                }
                for ($repeat = max(1, (int) $nodeCell->getAttribute('table:number-columns-repeated')); $repeat > 0; --$repeat) {
                    $table[$row][++$column] = $value;
                }
            }
        }

        return $table;
    }

    /**
     * The cells a range such as `local-table.$B$2:.$B$5` covers, row by row.
     *
     * @param array<int, array<int, null|string>> $table
     *
     * @return array<int, null|string>
     */
    protected function loadChartRange(array $table, string $address): array
    {
        if (1 !== preg_match('/\$?([A-Z]+)\$?(\d+)(?::[^$]*\$?([A-Z]+)\$?(\d+))?$/', $address, $matches)) {
            return [];
        }
        $columnFrom = $this->getChartColumnIndex($matches[1]);
        $columnTo = isset($matches[3]) ? $this->getChartColumnIndex($matches[3]) : $columnFrom;
        $rowTo = isset($matches[4]) ? (int) $matches[4] : (int) $matches[2];

        $cells = [];
        for ($row = (int) $matches[2]; $row <= $rowTo; ++$row) {
            for ($column = $columnFrom; $column <= $columnTo; ++$column) {
                $cells[] = $table[$row][$column] ?? null;
            }
        }

        return $cells;
    }

    /**
     * The index, from 1, of a spreadsheet column name (A => 1, Z => 26, AA => 27).
     */
    private function getChartColumnIndex(string $name): int
    {
        $index = 0;
        foreach (str_split($name) as $letter) {
            $index = $index * 26 + ord($letter) - 64;
        }

        return $index;
    }

    /**
     * The properties of one family the style of a chart element names, looked up in the automatic
     * styles of the chart's own document.
     */
    protected function getChartStyle(XMLReader $xmlReader, DOMElement $node, string $family): ?DOMElement
    {
        $styleName = $node->getAttribute('chart:style-name');
        if ('' === $styleName || false !== strpos($styleName, '"')) {
            return null;
        }

        return $xmlReader->getElement('/office:document-content/office:automatic-styles/style:style[@style:name="' . $styleName . '"]/style:' . $family . '-properties');
    }

    /**
     * The text of a title, its paragraphs one per line.
     */
    protected function loadChartText(XMLReader $xmlReader, DOMElement $nodeTitle): string
    {
        $text = [];
        foreach ($xmlReader->getElements('text:p', $nodeTitle) as $nodeParagraph) {
            $text[] = $nodeParagraph->textContent;
        }

        return implode("\n", $text);
    }

    /**
     * The font the style of a chart element names, or the default font where it names none.
     */
    protected function loadChartFont(XMLReader $xmlReader, DOMElement $node): Font
    {
        $textProps = $this->getChartStyle($xmlReader, $node, 'text');

        return $textProps instanceof DOMElement ? $this->loadFont($textProps) : new Font();
    }

    /**
     * The fill a chart graphic style names, where it names a colour and does not refuse a fill.
     */
    protected function loadChartFill(DOMElement $graphicProps): ?Fill
    {
        if (!$graphicProps->hasAttribute('draw:fill-color') || 'none' === $graphicProps->getAttribute('draw:fill')) {
            return null;
        }

        return (new Fill())->setFillType(Fill::FILL_SOLID)->setStartColor($this->loadChartColor($graphicProps, 'draw:fill-color'));
    }

    /**
     * The line a chart graphic style draws. An axis and a gridline count its width in points, a
     * serie and a data point in pixels, as the Writer writes them.
     */
    protected function loadChartOutline(DOMElement $graphicProps, bool $inPoints): Outline
    {
        $outline = new Outline();
        if ('none' === $graphicProps->getAttribute('draw:stroke')) {
            $outline->getFill()->setFillType(Fill::FILL_NONE);

            return $outline;
        }
        $outline->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($this->loadChartColor($graphicProps, 'svg:stroke-color'));
        if ($graphicProps->hasAttribute('svg:stroke-width')) {
            $width = (float) substr($graphicProps->getAttribute('svg:stroke-width'), 0, -2);
            $outline->setWidth((int) round($inPoints ? CommonDrawing::centimetersToPoints($width) : CommonDrawing::centimetersToPixels($width)));
        }

        return $outline;
    }

    protected function loadChartColor(DOMElement $graphicProps, string $attribute): Color
    {
        return new Color('FF' . strtoupper(substr($graphicProps->getAttribute($attribute), 1) ?: '000000'));
    }

    /**
     * Read where a shape sits, and the rotation a `draw:transform` gives it.
     *
     * A turned frame carries no `svg:x`/`svg:y`. It names a rotation about the origin and then a
     * translation, so the point written is where the top left corner lands once the shape has been
     * turned, and the offset has to be turned back out of it. ODF counts the angle the other way.
     */
    protected function loadShapeOffset(AbstractShape $shape, DOMElement $oNodeFrame): void
    {
        $pattern = '/rotate\s*\(\s*(-?[\d.]+)\s*\)\s*translate\s*\(\s*(-?[\d.]+)cm\s+(-?[\d.]+)cm\s*\)/';
        if (1 === preg_match($pattern, $oNodeFrame->getAttribute('draw:transform'), $matches)) {
            $rotation = -(float) $matches[1];
            $halfWidth = CommonDrawing::pixelsToCentimeters($shape->getWidth()) / 2;
            $halfHeight = CommonDrawing::pixelsToCentimeters($shape->getHeight()) / 2;
            $shape->setRotation((int) round(rad2deg($rotation)));
            $shape->setOffsetX((int) round(CommonDrawing::centimetersToPixels(
                (float) $matches[2] - $halfWidth + $halfWidth * cos($rotation) - $halfHeight * sin($rotation)
            )));
            $shape->setOffsetY((int) round(CommonDrawing::centimetersToPixels(
                (float) $matches[3] - $halfHeight + $halfWidth * sin($rotation) + $halfHeight * cos($rotation)
            )));

            return;
        }

        $shape->setOffsetX($oNodeFrame->hasAttribute('svg:x') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:x'), 0, -2)) : 0);
        $shape->setOffsetY($oNodeFrame->hasAttribute('svg:y') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:y'), 0, -2)) : 0);
    }

    /**
     * Read Shape RichText.
     *
     * @param null|ShapeContainerInterface $container where the shape goes, the slide itself unless
     *                                                the frame was read out of the slide note
     */
    protected function loadShapeRichText(DOMElement $oNodeFrame, ?ShapeContainerInterface $container = null): void
    {
        // Core
        $oShape = new RichText();
        ($container ?? $this->oPhpPresentation->getActiveSlide())->addShape($oShape);
        $oShape->setParagraphs([]);

        $oShape->setDescription($this->loadShapeDescription($oNodeFrame));
        $oShape->setDecorative($this->loadShapeDecorative($oNodeFrame));
        $oShape->setWidth($oNodeFrame->hasAttribute('svg:width') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:width'), 0, -2)) : 0);
        $oShape->setHeight($oNodeFrame->hasAttribute('svg:height') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:height'), 0, -2)) : 0);
        $this->loadShapeOffset($oShape, $oNodeFrame);

        if ($oNodeFrame->hasAttribute('draw:style-name')) {
            $keyStyle = $oNodeFrame->getAttribute('draw:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                if (null !== $this->arrayStyles[$keyStyle]['columns']) {
                    $oShape->setColumns($this->arrayStyles[$keyStyle]['columns']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['columnSpacing']) {
                    $oShape->setColumnSpacing($this->arrayStyles[$keyStyle]['columnSpacing']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['columnsRTL']) {
                    $oShape->setColumnsRTL($this->arrayStyles[$keyStyle]['columnsRTL']);
                }
                $this->applyShapeBorder($oShape, $this->arrayStyles[$keyStyle]['border']);
                // the graphic style of a text box carries its fill and its shadow just as the one
                // of a drawing does, and both were read out of it and then only handed to a drawing
                if (null !== $this->arrayStyles[$keyStyle]['fill']) {
                    $oShape->setFill($this->arrayStyles[$keyStyle]['fill']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['shadow']) {
                    $oShape->setShadow($this->arrayStyles[$keyStyle]['shadow']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['insetBottom']) {
                    $oShape->setInsetBottom($this->arrayStyles[$keyStyle]['insetBottom']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['insetLeft']) {
                    $oShape->setInsetLeft($this->arrayStyles[$keyStyle]['insetLeft']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['insetRight']) {
                    $oShape->setInsetRight($this->arrayStyles[$keyStyle]['insetRight']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['insetTop']) {
                    $oShape->setInsetTop($this->arrayStyles[$keyStyle]['insetTop']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['verticalAlignCenter']) {
                    $oShape->setVerticalAlignCenter($this->arrayStyles[$keyStyle]['verticalAlignCenter']);
                }
                if (null !== $this->arrayStyles[$keyStyle]['wrap']) {
                    $oShape->setWrap($this->arrayStyles[$keyStyle]['wrap']);
                }
            }
        }

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
    /**
     * @param Cell|RichText $oShape anything that opens a paragraph: a text shape or a table cell
     */
    protected function readParagraph($oShape, DOMElement $oNodeParent): void
    {
        $oParagraph = $oShape->createParagraph();
        if ($oNodeParent->hasAttribute('text:style-name')) {
            $keyStyle = $oNodeParent->getAttribute('text:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                if (null !== $this->arrayStyles[$keyStyle]['alignment']) {
                    $oParagraph->setAlignment($this->arrayStyles[$keyStyle]['alignment']);
                }
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
                    $oParagraph->setLineSpacing((int) $this->arrayStyles[$keyStyle]['lineSpacing']);
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
            // A field is a run whose text the reading application recomputes, and OpenDocument
            // wraps it in the same `text:span` an ordinary run gets, so what it is has to be
            // looked for before the run is made
            $oNodeField = $this->oXMLReader->getElement(
                '(' . implode('|', array_keys(self::FIELD_OOXML)) . ')',
                $oNodeParent
            );
            $oTextRun = $oNodeField instanceof DOMElement
                ? $oParagraph->createField(self::FIELD_OOXML[$oNodeField->nodeName])
                : $oParagraph->createTextRun();
            if ($oNodeParent->hasAttribute('text:style-name')) {
                $keyStyle = $oNodeParent->getAttribute('text:style-name');
                if (isset($this->arrayStyles[$keyStyle])) {
                    $oTextRun->setFont($this->arrayStyles[$keyStyle]['font']);
                    if (null !== $this->arrayStyles[$keyStyle]['language']) {
                        $oTextRun->setLanguage($this->arrayStyles[$keyStyle]['language']);
                    }
                }
            }
            $oTextRunLink = $this->oXMLReader->getElement('text:a', $oNodeParent);
            if ($oTextRunLink instanceof DOMElement) {
                $oTextRun->setText($oTextRunLink->nodeValue);
                if ($oTextRunLink->hasAttribute('xlink:href')) {
                    $href = $oTextRunLink->getAttribute('xlink:href');
                    if (0 === strpos($href, '#') && isset($this->arraySlideNumbers[substr($href, 1)])) {
                        $oTextRun->getHyperlink()->setSlideNumber($this->arraySlideNumbers[substr($href, 1)]);
                    } else {
                        $oTextRun->getHyperlink()->setUrl($href);
                    }
                }
            } elseif ($oNodeField instanceof DOMElement) {
                // the span holds the field, and the field holds the text it stands in for
                $oTextRun->setText($oNodeField->nodeValue);
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
    /**
     * Read Shape Table.
     *
     * An ODF table lives inside a `draw:frame` like any other shape, and says which of its rows
     * are styled apart on the table itself rather than by where they sit.
     */
    protected function loadShapeTable(DOMElement $oNodeFrame): void
    {
        $columns = 0;
        foreach ($this->oXMLReader->getElements('table:table/table:table-column', $oNodeFrame) as $oNodeColumn) {
            $columns += $oNodeColumn instanceof DOMElement && $oNodeColumn->hasAttribute('table:number-columns-repeated')
                ? (int) $oNodeColumn->getAttribute('table:number-columns-repeated')
                : 1;
        }

        $oShape = $this->oPhpPresentation->getActiveSlide()->createTableShape(max($columns, 1));
        $oShape->setDescription($this->loadShapeDescription($oNodeFrame));
        $oShape->setDecorative($this->loadShapeDecorative($oNodeFrame));
        $oShape->setWidth($oNodeFrame->hasAttribute('svg:width') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:width'), 0, -2)) : 0);
        $oShape->setHeight($oNodeFrame->hasAttribute('svg:height') ? CommonDrawing::centimetersToPixels((float) substr($oNodeFrame->getAttribute('svg:height'), 0, -2)) : 0);
        $this->loadShapeOffset($oShape, $oNodeFrame);

        // A drawing table says which of its rows are styled apart with these two flags, which are
        // what `firstRow` and `bandRow` say on `a:tblPr`
        $oNodeTable = $this->oXMLReader->getElement('table:table', $oNodeFrame);
        if ($oNodeTable instanceof DOMElement) {
            $oShape->setFirstRow('true' === $oNodeTable->getAttribute('table:use-first-row-styles'));
            if ($oNodeTable->hasAttribute('table:use-banding-rows-styles')) {
                $oShape->setBandRow('true' === $oNodeTable->getAttribute('table:use-banding-rows-styles'));
            }
        }

        foreach ($this->oXMLReader->getElements('table:table/table:table-row', $oNodeFrame) as $oNodeRow) {
            if ($oNodeRow instanceof DOMElement) {
                $this->loadTableRow($oShape->createRow(), $oNodeRow);
            }
        }
    }

    /**
     * Read one row of a table, and every cell it holds.
     */
    protected function loadTableRow(Row $oRow, DOMElement $oNodeRow): void
    {
        $rowStyle = $this->getStyle($oNodeRow, 'table:style-name');
        if (null !== $rowStyle['rowHeight']) {
            $oRow->setHeight($rowStyle['rowHeight']);
        }

        // A cell that names no style of its own takes the one the row names for all of them
        $defaultCellStyle = $oNodeRow->hasAttribute('table:default-cell-style-name')
            ? $oNodeRow->getAttribute('table:default-cell-style-name')
            : '';

        $cellIndex = 0;
        foreach ($this->oXMLReader->getElements('table:table-cell', $oNodeRow) as $oNodeCell) {
            if (!$oNodeCell instanceof DOMElement || !$oRow->hasCell($cellIndex)) {
                continue;
            }
            $this->loadTableCell($oRow->getCell($cellIndex), $oNodeCell, $defaultCellStyle);
            ++$cellIndex;
        }
    }

    /**
     * Read one cell of a table: the fill and the borders its style names, and its text.
     */
    protected function loadTableCell(Cell $oCell, DOMElement $oNodeCell, string $defaultCellStyle = ''): void
    {
        $cellStyle = $this->getStyle($oNodeCell, 'table:style-name', $defaultCellStyle);
        if (null !== $cellStyle['fill']) {
            $oCell->setFill($cellStyle['fill']);
        }
        if (null !== $cellStyle['borders']) {
            $oCell->setBorders($cellStyle['borders']);
        }

        // A cell holds its text in `text:p`, which is what this Writer and LibreOffice both put
        // there
        $oCell->setParagraphs([]);
        foreach ($this->oXMLReader->getElements('text:p', $oNodeCell) as $oNodeParagraph) {
            if ($oNodeParagraph instanceof DOMElement) {
                $this->levelParagraph = 0;
                $this->readParagraph($oCell, $oNodeParagraph);
            }
        }

        if (count($oCell->getParagraphs()) > 0) {
            $oCell->setActiveParagraph(0);
        }
    }

    /**
     * The style a node names, with every key the reader parses present.
     *
     * @return array<string, mixed>
     */
    protected function getStyle(DOMElement $oNode, string $attribute, string $fallback = ''): array
    {
        $keyStyle = $oNode->hasAttribute($attribute) ? $oNode->getAttribute($attribute) : $fallback;

        return $this->arrayStyles[$keyStyle] ?? array_fill_keys(self::STYLE_KEYS, null);
    }

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

    /**
     * The raise `style:text-position` names, in the thousandths of a percent a baseline is held in.
     * Its `super` and `sub` answer with the values PowerPoint writes.
     */
    protected function baselineFromTextPosition(string $value): int
    {
        $position = strtok(trim($value), " \t") ?: '';

        if ('super' === $position) {
            return Font::BASELINE_SUPERSCRIPT;
        }
        if ('sub' === $position) {
            return Font::BASELINE_SUBSCRIPT;
        }

        return (int) round(((float) rtrim($position, '%')) * 1000);
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
