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

namespace PhpOffice\PhpPresentation\Reader;

use ZipArchive;
use PhpOffice\Common\XMLReader;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Style\Alignment;

/**
 * Serialized format reader
 */
class ODPresentation implements ReaderInterface
{
    /**
     * Output Object
     * @var PhpPresentation
     */
    protected $oPhpPresentation;
    /**
     * Output Object
     * @var \ZipArchive
     */
    protected $oZip;
    /**
     * @var array[]
     */
    protected $arrayStyles = array();
    /**
     * @var array[]
     */
    protected $arrayCommonStyles = array();
    /**
     * @var \PhpOffice\Common\XMLReader
     */
    protected $oXMLReader;

    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     *
     * @param  string $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function canRead($pFilename)
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     *
     * @param  string    $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function fileSupportsUnserializePhpPresentation($pFilename = '')
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }
        
        $oZip = new ZipArchive();
        // Is it a zip ?
        if ($oZip->open($pFilename) === true) {
            // Is it an OpenXML Document ?
            // Is it a Presentation ?
            if (is_array($oZip->statName('META-INF/manifest.xml')) && is_array($oZip->statName('mimetype')) && $oZip->getFromName('mimetype') == 'application/vnd.oasis.opendocument.presentation') {
                return true;
            }
        }
        return false;
    }

    /**
     * Loads PhpPresentation Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     * @throws \Exception
     */
    public function load($pFilename)
    {
        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPresentation\Reader\ODPresentation: " . $pFilename . ".");
        }

        return $this->loadFile($pFilename);
    }

    /**
     * Load PhpPresentation Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     */
    protected function loadFile($pFilename)
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();
        
        $this->oZip = new ZipArchive();
        $this->oZip->open($pFilename);
        
        $this->oXMLReader = new XMLReader();
        if ($this->oXMLReader->getDomFromZip($pFilename, 'meta.xml') !== false) {
            $this->loadDocumentProperties();
        }
        $this->oXMLReader = new XMLReader();
        if ($this->oXMLReader->getDomFromZip($pFilename, 'styles.xml') !== false) {
            $this->loadStylesFile();
        }
        $this->oXMLReader = new XMLReader();
        if ($this->oXMLReader->getDomFromZip($pFilename, 'content.xml') !== false) {
            $this->loadSlides();
        }

        return $this->oPhpPresentation;
    }
    
    /**
     * Read Document Properties
     */
    protected function loadDocumentProperties()
    {
        $arrayProperties = array(
            '/office:document-meta/office:meta/meta:initial-creator' => 'setCreator',
            '/office:document-meta/office:meta/dc:creator' => 'setLastModifiedBy',
            '/office:document-meta/office:meta/dc:title' => 'setTitle',
            '/office:document-meta/office:meta/dc:description' => 'setDescription',
            '/office:document-meta/office:meta/dc:subject' => 'setSubject',
            '/office:document-meta/office:meta/meta:keyword' => 'setKeywords',
            '/office:document-meta/office:meta/meta:creation-date' => 'setCreated',
            '/office:document-meta/office:meta/dc:date' => 'setModified',
        );
        $oProperties = $this->oPhpPresentation->getProperties();
        foreach ($arrayProperties as $path => $property) {
            if (is_object($oElement = $this->oXMLReader->getElement($path))) {
                if (in_array($property, array('setCreated', 'setModified'))) {
                    $oDateTime = new \DateTime();
                    $oDateTime->createFromFormat(\DateTime::W3C, $oElement->nodeValue);
                    $oProperties->{$property}($oDateTime->getTimestamp());
                } else {
                    $oProperties->{$property}($oElement->nodeValue);
                }
            }
        }
    }
    
    /**
     * Extract all slides
     */
    protected function loadSlides()
    {
        foreach ($this->oXMLReader->getElements('/office:document-content/office:automatic-styles/*') as $oElement) {
            if ($oElement->hasAttribute('style:name')) {
                $this->loadStyle($oElement);
            }
        }
        foreach ($this->oXMLReader->getElements('/office:document-content/office:body/office:presentation/draw:page') as $oElement) {
            if ($oElement->nodeName == 'draw:page') {
                $this->loadSlide($oElement);
            }
        }
    }
    
    /**
     * Extract style
     * @param \DOMElement $nodeStyle
     */
    protected function loadStyle(\DOMElement $nodeStyle)
    {
        $keyStyle = $nodeStyle->getAttribute('style:name');

        $nodeDrawingPageProps = $this->oXMLReader->getElement('style:drawing-page-properties', $nodeStyle);
        if ($nodeDrawingPageProps) {
            // Read Background Color
            if ($nodeDrawingPageProps->hasAttribute('draw:fill-color') && $nodeDrawingPageProps->getAttribute('draw:fill') == 'solid') {
                $oBackground = new \PhpOffice\PhpPresentation\Slide\Background\Color();
                $oColor = new Color();
                $oColor->setRGB(substr($nodeDrawingPageProps->getAttribute('draw:fill-color'), -6));
                $oBackground->setColor($oColor);
            }
            // Read Background Image
            if ($nodeDrawingPageProps->getAttribute('draw:fill') == 'bitmap' && $nodeDrawingPageProps->hasAttribute('draw:fill-image-name')) {
                $nameStyle = $nodeDrawingPageProps->getAttribute('draw:fill-image-name');
                if (!empty($this->arrayCommonStyles[$nameStyle]) && $this->arrayCommonStyles[$nameStyle]['type'] == 'image' && !empty($this->arrayCommonStyles[$nameStyle]['path'])) {
                    $tmpBkgImg = tempnam(sys_get_temp_dir(), 'PhpPresentationReaderODPBkg');
                    $contentImg = $this->oZip->getFromName($this->arrayCommonStyles[$nameStyle]['path']);
                    file_put_contents($tmpBkgImg, $contentImg);

                    $oBackground = new \PhpOffice\PhpPresentation\Slide\Background\Image();
                    $oBackground->setPath($tmpBkgImg);
                }
            }
        }

        $nodeGraphicProps = $this->oXMLReader->getElement('style:graphic-properties', $nodeStyle);
        if ($nodeGraphicProps) {
            // Read Shadow
            if ($nodeGraphicProps->hasAttribute('draw:shadow') && $nodeGraphicProps->getAttribute('draw:shadow') == 'visible') {
                $oShadow = new Shadow();
                $oShadow->setVisible(true);
                if ($nodeGraphicProps->hasAttribute('draw:shadow-color')) {
                    $oShadow->getColor()->setRGB(substr($nodeGraphicProps->getAttribute('draw:shadow-color'), -6));
                }
                if ($nodeGraphicProps->hasAttribute('draw:shadow-opacity')) {
                    $oShadow->setAlpha(100 - (int)substr($nodeGraphicProps->getAttribute('draw:shadow-opacity'), 0, -1));
                }
                if ($nodeGraphicProps->hasAttribute('draw:shadow-offset-x') && $nodeGraphicProps->hasAttribute('draw:shadow-offset-y')) {
                    $offsetX = substr($nodeGraphicProps->getAttribute('draw:shadow-offset-x'), 0, -2);
                    $offsetY = substr($nodeGraphicProps->getAttribute('draw:shadow-offset-y'), 0, -2);
                    $distance = 0;
                    if ($offsetX != 0) {
                        $distance = ($offsetX < 0 ? $offsetX * -1 : $offsetX);
                    } elseif ($offsetY != 0) {
                        $distance = ($offsetY < 0 ? $offsetY * -1 : $offsetY);
                    }
                    $oShadow->setDirection(rad2deg(atan2($offsetY, $offsetX)));
                    $oShadow->setDistance(CommonDrawing::centimetersToPixels($distance));
                }
            }
        }
        
        $nodeTextProperties = $this->oXMLReader->getElement('style:text-properties', $nodeStyle);
        if ($nodeTextProperties) {
            $oFont = new Font();
            if ($nodeTextProperties->hasAttribute('fo:color')) {
                $oFont->getColor()->setRGB(substr($nodeTextProperties->getAttribute('fo:color'), -6));
            }
            if ($nodeTextProperties->hasAttribute('fo:font-family')) {
                $oFont->setName($nodeTextProperties->getAttribute('fo:font-family'));
            }
            if ($nodeTextProperties->hasAttribute('fo:font-weight') && $nodeTextProperties->getAttribute('fo:font-weight') == 'bold') {
                $oFont->setBold(true);
            }
            if ($nodeTextProperties->hasAttribute('fo:font-size')) {
                $oFont->setSize(substr($nodeTextProperties->getAttribute('fo:font-size'), 0, -2));
            }
        }

        $nodeParagraphProps = $this->oXMLReader->getElement('style:paragraph-properties', $nodeStyle);
        if ($nodeParagraphProps) {
            $oAlignment = new Alignment();
            if ($nodeParagraphProps->hasAttribute('fo:text-align')) {
                $oAlignment->setHorizontal($nodeParagraphProps->getAttribute('fo:text-align'));
            }
        }
        
        if ($nodeStyle->nodeName == 'text:list-style') {
            $arrayListStyle = array();
            foreach ($this->oXMLReader->getElements('text:list-level-style-bullet', $nodeStyle) as $oNodeListLevel) {
                $oAlignment = new Alignment();
                $oBullet = new Bullet();
                $oBullet->setBulletType(Bullet::TYPE_NONE);
                if ($oNodeListLevel->hasAttribute('text:level')) {
                    $oAlignment->setLevel((int) $oNodeListLevel->getAttribute('text:level') - 1);
                }
                if ($oNodeListLevel->hasAttribute('text:bullet-char')) {
                    $oBullet->setBulletChar($oNodeListLevel->getAttribute('text:bullet-char'));
                    $oBullet->setBulletType(Bullet::TYPE_BULLET);
                }
                
                $oNodeListProperties = $this->oXMLReader->getElement('style:list-level-properties', $oNodeListLevel);
                if ($oNodeListProperties) {
                    if ($oNodeListProperties->hasAttribute('text:min-label-width')) {
                        $oAlignment->setIndent((int)round(CommonDrawing::centimetersToPixels(substr($oNodeListProperties->getAttribute('text:min-label-width'), 0, -2))));
                    }
                    if ($oNodeListProperties->hasAttribute('text:space-before')) {
                        $iSpaceBefore = CommonDrawing::centimetersToPixels(substr($oNodeListProperties->getAttribute('text:space-before'), 0, -2));
                        $iMarginLeft = $iSpaceBefore + $oAlignment->getIndent();
                        $oAlignment->setMarginLeft($iMarginLeft);
                    }
                }
                $oNodeTextProperties = $this->oXMLReader->getElement('style:text-properties', $oNodeListLevel);
                if ($oNodeTextProperties) {
                    if ($oNodeTextProperties->hasAttribute('fo:font-family')) {
                        $oBullet->setBulletFont($oNodeTextProperties->getAttribute('fo:font-family'));
                    }
                }
                
                $arrayListStyle[$oAlignment->getLevel()] = array(
                    'alignment' => $oAlignment,
                    'bullet' => $oBullet,
                );
            }
        }
        
        $this->arrayStyles[$keyStyle] = array(
            'alignment' => isset($oAlignment) ? $oAlignment : null,
            'background' => isset($oBackground) ? $oBackground : null,
            'font' => isset($oFont) ? $oFont : null,
            'shadow' => isset($oShadow) ? $oShadow : null,
            'listStyle' => isset($arrayListStyle) ? $arrayListStyle : null,
        );
        
        return true;
    }

    /**
     * Read Slide
     *
     * @param \DOMElement $nodeSlide
     */
    protected function loadSlide(\DOMElement $nodeSlide)
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
            if ($this->oXMLReader->getElement('draw:image', $oNodeFrame)) {
                $this->loadShapeDrawing($oNodeFrame);
                continue;
            }
            if ($this->oXMLReader->getElement('draw:text-box', $oNodeFrame)) {
                $this->loadShapeRichText($oNodeFrame);
                continue;
            }
        }
        return true;
    }
    
    /**
     * Read Shape Drawing
     *
     * @param \DOMElement $oNodeFrame
     */
    protected function loadShapeDrawing(\DOMElement $oNodeFrame)
    {
        // Core
        $oShape = new Gd();
        $oShape->getShadow()->setVisible(false);

        $oNodeImage = $this->oXMLReader->getElement('draw:image', $oNodeFrame);
        if ($oNodeImage) {
            if ($oNodeImage->hasAttribute('xlink:href')) {
                $sFilename = $oNodeImage->getAttribute('xlink:href');
                // svm = StarView Metafile
                if (pathinfo($sFilename, PATHINFO_EXTENSION) == 'svm') {
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
        $oShape->setWidth($oNodeFrame->hasAttribute('svg:width') ? (int)round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:width'), 0, -2))) : '');
        $oShape->setHeight($oNodeFrame->hasAttribute('svg:height') ? (int)round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:height'), 0, -2))) : '');
        $oShape->setResizeProportional(true);
        $oShape->setOffsetX($oNodeFrame->hasAttribute('svg:x') ? (int)round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:x'), 0, -2))) : '');
        $oShape->setOffsetY($oNodeFrame->hasAttribute('svg:y') ? (int)round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:y'), 0, -2))) : '');
        
        if ($oNodeFrame->hasAttribute('draw:style-name')) {
            $keyStyle = $oNodeFrame->getAttribute('draw:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                $oShape->setShadow($this->arrayStyles[$keyStyle]['shadow']);
            }
        }
        
        $this->oPhpPresentation->getActiveSlide()->addShape($oShape);
    }

    /**
     * Read Shape RichText
     *
     * @param \DOMElement $oNodeFrame
     */
    protected function loadShapeRichText(\DOMElement $oNodeFrame)
    {
        // Core
        $oShape = $this->oPhpPresentation->getActiveSlide()->createRichTextShape();
        $oShape->setParagraphs(array());
        
        $oShape->setWidth($oNodeFrame->hasAttribute('svg:width') ? (int)round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:width'), 0, -2))) : '');
        $oShape->setHeight($oNodeFrame->hasAttribute('svg:height') ? (int)round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:height'), 0, -2))) : '');
        $oShape->setOffsetX($oNodeFrame->hasAttribute('svg:x') ? (int)round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:x'), 0, -2))) : '');
        $oShape->setOffsetY($oNodeFrame->hasAttribute('svg:y') ? (int)round(CommonDrawing::centimetersToPixels(substr($oNodeFrame->getAttribute('svg:y'), 0, -2))) : '');
        
        foreach ($this->oXMLReader->getElements('draw:text-box/*', $oNodeFrame) as $oNodeParagraph) {
            $this->levelParagraph = 0;
            if ($oNodeParagraph->nodeName == 'text:p') {
                $this->readParagraph($oShape, $oNodeParagraph);
            }
            if ($oNodeParagraph->nodeName == 'text:list') {
                $this->readList($oShape, $oNodeParagraph);
            }
        }
        
        if (count($oShape->getParagraphs()) > 0) {
            $oShape->setActiveParagraph(0);
        }
    }
    
    protected $levelParagraph = 0;
    
    /**
     * Read Paragraph
     * @param RichText $oShape
     * @param \DOMElement $oNodeParent
     */
    protected function readParagraph(RichText $oShape, \DOMElement $oNodeParent)
    {
        $oParagraph = $oShape->createParagraph();
        $oDomList = $this->oXMLReader->getElements('text:span', $oNodeParent);
        if ($oDomList->length == 0) {
            $this->readParagraphItem($oParagraph, $oNodeParent);
        } else {
            foreach ($oDomList as $oNodeRichTextElement) {
                $this->readParagraphItem($oParagraph, $oNodeRichTextElement);
            }
        }
    }
    
    /**
     * Read Paragraph Item
     * @param RichText $oShape
     * @param \DOMElement $oNodeParent
     */
    protected function readParagraphItem(Paragraph $oParagraph, \DOMElement $oNodeParent)
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
            if ($oTextRunLink = $this->oXMLReader->getElement('text:a', $oNodeParent)) {
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
     * Read List
     *
     * @param RichText $oShape
     * @param \DOMElement $oNodeParent
     */
    protected function readList(RichText $oShape, \DOMElement $oNodeParent)
    {
        foreach ($this->oXMLReader->getElements('text:list-item/*', $oNodeParent) as $oNodeListItem) {
            if ($oNodeListItem->nodeName == 'text:p') {
                $this->readListItem($oShape, $oNodeListItem, $oNodeParent);
            }
            if ($oNodeListItem->nodeName == 'text:list') {
                $this->levelParagraph++;
                $this->readList($oShape, $oNodeListItem);
                $this->levelParagraph--;
            }
        }
    }
    
    /**
     * Read List Item
     * @param RichText $oShape
     * @param \DOMElement $oNodeParent
     * @param \DOMElement $oNodeParagraph
     */
    protected function readListItem(RichText $oShape, \DOMElement $oNodeParent, \DOMElement $oNodeParagraph)
    {
        $oParagraph = $oShape->createParagraph();
        if ($oNodeParagraph->hasAttribute('text:style-name')) {
            $keyStyle = $oNodeParagraph->getAttribute('text:style-name');
            if (isset($this->arrayStyles[$keyStyle])) {
                $oParagraph->setAlignment($this->arrayStyles[$keyStyle]['listStyle'][$this->levelParagraph]['alignment']);
                $oParagraph->setBulletStyle($this->arrayStyles[$keyStyle]['listStyle'][$this->levelParagraph]['bullet']);
            }
        }
        foreach ($this->oXMLReader->getElements('text:span', $oNodeParent) as $oNodeRichTextElement) {
            $this->readParagraphItem($oParagraph, $oNodeRichTextElement);
        }
    }

    /**
     * Load file 'styles.xml'
     */
    protected function loadStylesFile()
    {
        foreach ($this->oXMLReader->getElements('/office:document-styles/office:styles/*') as $oElement) {
            if ($oElement->nodeName == 'draw:fill-image') {
                $this->arrayCommonStyles[$oElement->getAttribute('draw:name')] = array(
                    'type' => 'image',
                    'path' => $oElement->hasAttribute('xlink:href') ? $oElement->getAttribute('xlink:href') : null
                );
            }
        }
    }
}
