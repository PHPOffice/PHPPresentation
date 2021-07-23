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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Reader;

use DateTime;
use DOMElement;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLReader;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
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
     * @var \ZipArchive
     */
    protected $oZip;
    /**
     * @var array[]
     */
    protected $arrayStyles = [];
    /**
     * @var array[]
     */
    protected $arrayCommonStyles = [];
    /**
     * @var \PhpOffice\Common\XMLReader
     */
    protected $oXMLReader;
    /**
     * @var int
     */
    protected $levelParagraph = 0;

    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     *
     * @throws \Exception
     */
    public function canRead(string $pFilename): bool
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     *
     * @throws \Exception
     */
    public function fileSupportsUnserializePhpPresentation(string $pFilename = ''): bool
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception('Could not open ' . $pFilename . ' for reading! File does not exist.');
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
     *
     * @throws \Exception
     */
    public function load(string $pFilename): PhpPresentation
    {
        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPresentation\Reader\ODPresentation: " . $pFilename . '.');
        }

        return $this->loadFile($pFilename);
    }

    /**
     * Load PhpPresentation Serialized file.
     *
     * @param string $pFilename
     *
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     *
     * @throws \Exception
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
        $oProperties = $this->oPhpPresentation->getDocumentProperties();
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
                $oProperties->{$property}($value);
            }
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
     *
     * @return bool
     *
     * @throws \Exception
     */
    protected function loadStyle(DOMElement $nodeStyle)
    {
        $keyStyle = $nodeStyle->getAttribute('style:name');

        $nodeDrawingPageProps = $this->oXMLReader->getElement('style:drawing-page-properties', $nodeStyle);
        if ($nodeDrawingPageProps instanceof DOMElement) {
            // Read Background Color
            if ($nodeDrawingPageProps->hasAttribute('draw:fill-color') && 'solid' == $nodeDrawingPageProps->getAttribute('draw:fill')) {
                $oBackground = new \PhpOffice\PhpPresentation\Slide\Background\Color();
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
                    $oShadow->setDistance((int) round(CommonDrawing::centimetersToPixels($distance)));
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
                            $oAlignment->setIndent((int) round(CommonDrawing::centimetersToPixels(substr($oNodeListProperties->getAttribute('text:min-label-width'), 0, -2))));
                        }
                        if ($oNodeListProperties->hasAttribute('text:space-before')) {
                            $iSpaceBefore = (int) CommonDrawing::centimetersToPixels(substr($oNodeListProperties->getAttribute('text:space-before'), 0, -2));
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
            'alignment' => isset($oAlignment) ? $oAlignment : null,
            'background' => isset($oBackground) ? $oBackground : null,
            'fill' => isset($oFill) ? $oFill : null,
            'font' => isset($oFont) ? $oFont : null,
            'shadow' => isset($oShadow) ? $oShadow : null,
            'listStyle' => isset($arrayListStyle) ? $arrayListStyle : null,
        ];

        return true;
    }

    /**
     * Read Slide.
     *
     * @throws \Exception
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
                if ($this->oXMLReader->getElement('draw:image', $oNodeFrame)) {
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
     *
     * @throws \Exception
     */
    protected function loadShapeDrawing(DOMElement $oNodeFrame): void
    {
        // Core
        $oShape = new Gd();
        $oShape->getShadow()->setVisible(false);

        $oNodeImage = $this->oXMLReader->getElement('draw:image', $oNodeFrame);
        if ($oNodeImage instanceof DOMElement) {
            if ($oNodeImage->hasAttribute('xlink:href')) {
                $sFilename = $oNodeImage->getAttribute('xlink:href');
                // svm = StarView Metafile
                if ('svm' == pathinfo($sFilename, PATHINFO_EXTENSION)) {
                    return;
                }
                $imageFile = $this->oZip->getFromName($sFilename);
                if (!empty($imageFile)) {
                    $oShape->setImageResource(imagecreatefromstring($imageFile));
                }
            }
        }

        $oShape->setName($oNodeFrame->hasAttribute('draw:name') ? $oNodeFrame->getAttribute('draw:name') : '');
        $oShape->setDescription($oNodeFrame->hasAttribute('draw:name') ? $oNodeFrame->getAttribute('draw:name') : '');
        $oShape->setResizeProportional(false);
        $oShape->setWidth($oNodeFrame->hasAttribute('svg:width') ? (int) round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:width'), 0, -2))) : '');
        $oShape->setHeight($oNodeFrame->hasAttribute('svg:height') ? (int) round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:height'), 0, -2))) : '');
        $oShape->setResizeProportional(true);
        $oShape->setOffsetX($oNodeFrame->hasAttribute('svg:x') ? (int) round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:x'), 0, -2))) : '');
        $oShape->setOffsetY($oNodeFrame->hasAttribute('svg:y') ? (int) round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:y'), 0, -2))) : '');

        if ($oNodeFrame->hasAttribute('draw:style-name')) {
            $keyStyle = $oNodeFrame->getAttribute('draw:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                $oShape->setShadow($this->arrayStyles[$keyStyle]['shadow']);
                $oShape->setFill($this->arrayStyles[$keyStyle]['fill']);
            }
        }

        $this->oPhpPresentation->getActiveSlide()->addShape($oShape);
    }

    /**
     * Read Shape RichText.
     *
     * @throws \Exception
     */
    protected function loadShapeRichText(DOMElement $oNodeFrame): void
    {
        // Core
        $oShape = $this->oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->setParagraphs([]);

        $oShape->setWidth($oNodeFrame->hasAttribute('svg:width') ? (int) round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:width'), 0, -2))) : '');
        $oShape->setHeight($oNodeFrame->hasAttribute('svg:height') ? (int) round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:height'), 0, -2))) : '');
        $oShape->setOffsetX($oNodeFrame->hasAttribute('svg:x') ? (int) round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:x'), 0, -2))) : '');
        $oShape->setOffsetY($oNodeFrame->hasAttribute('svg:y') ? (int) round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:y'), 0, -2))) : '');

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
     *
     * @throws \Exception
     */
    protected function readParagraph(RichText $oShape, DOMElement $oNodeParent): void
    {
        $oParagraph = $oShape->createParagraph();
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
     *
     * @throws \Exception
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
     *
     * @throws \Exception
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
     *
     * @throws \Exception
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
}
