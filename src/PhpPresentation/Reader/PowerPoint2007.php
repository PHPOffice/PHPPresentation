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

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Slide\AbstractSlide;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use PhpOffice\PhpPresentation\Style\SchemeColor;
use PhpOffice\PhpPresentation\Style\TextStyle;
use ZipArchive;
use PhpOffice\Common\XMLReader;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;

/**
 * Serialized format reader
 */
class PowerPoint2007 implements ReaderInterface
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
    protected $arrayRels = array();
    /**
     * @var SlideLayout[]
     */
    protected $arraySlideLayouts = array();
    /*
     * @var string
     */
    protected $filename;

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
     * @param  string $pFilename
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
            if (is_array($oZip->statName('[Content_Types].xml')) && is_array($oZip->statName('ppt/presentation.xml'))) {
                return true;
            }
        }
        return false;
    }

    /**
     * Loads PhpPresentation Serialized file
     *
     * @param  string $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     * @throws \Exception
     */
    public function load($pFilename)
    {
        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPresentation\Reader\PowerPoint2007: " . $pFilename . ".");
        }

        return $this->loadFile($pFilename);
    }

    /**
     * Load PhpPresentation Serialized file
     *
     * @param  string $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     */
    protected function loadFile($pFilename)
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();
        $this->oPhpPresentation->setAllMasterSlides(array());
        $this->filename = $pFilename;

        $this->oZip = new ZipArchive();
        $this->oZip->open($this->filename);
        $docPropsCore = $this->oZip->getFromName('docProps/core.xml');
        if ($docPropsCore !== false) {
            $this->loadDocumentProperties($docPropsCore);
        }

        $docPropsCustom = $this->oZip->getFromName('docProps/custom.xml');
        if ($docPropsCustom !== false) {
            $this->loadCustomProperties($docPropsCustom);
        }

        $pptViewProps = $this->oZip->getFromName('ppt/viewProps.xml');
        if ($pptViewProps !== false) {
            $this->loadViewProperties($pptViewProps);
        }

        $pptPresentation = $this->oZip->getFromName('ppt/presentation.xml');
        if ($pptPresentation !== false) {
            $this->loadDocumentLayout($pptPresentation);
            $this->loadSlides($pptPresentation);
        }

        return $this->oPhpPresentation;
    }

    /**
     * Read Document Layout
     * @param $sPart
     */
    protected function loadDocumentLayout($sPart)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            foreach ($xmlReader->getElements('/p:presentation/p:sldSz') as $oElement) {
                $type = $oElement->getAttribute('type');
                $oLayout = $this->oPhpPresentation->getLayout();
                if ($type == DocumentLayout::LAYOUT_CUSTOM) {
                    $oLayout->setCX($oElement->getAttribute('cx'));
                    $oLayout->setCY($oElement->getAttribute('cy'));
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
     * Read Document Properties
     * @param string $sPart
     */
    protected function loadDocumentProperties($sPart)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            $arrayProperties = array(
                '/cp:coreProperties/dc:creator' => 'setCreator',
                '/cp:coreProperties/cp:lastModifiedBy' => 'setLastModifiedBy',
                '/cp:coreProperties/dc:title' => 'setTitle',
                '/cp:coreProperties/dc:description' => 'setDescription',
                '/cp:coreProperties/dc:subject' => 'setSubject',
                '/cp:coreProperties/cp:keywords' => 'setKeywords',
                '/cp:coreProperties/cp:category' => 'setCategory',
                '/cp:coreProperties/dcterms:created' => 'setCreated',
                '/cp:coreProperties/dcterms:modified' => 'setModified',
            );
            $oProperties = $this->oPhpPresentation->getDocumentProperties();
            foreach ($arrayProperties as $path => $property) {
                if (is_object($oElement = $xmlReader->getElement($path))) {
                    if ($oElement->hasAttribute('xsi:type') && $oElement->getAttribute('xsi:type') == 'dcterms:W3CDTF') {
                        $oDateTime = new \DateTime();
                        $oDateTime->createFromFormat(\DateTime::W3C, $oElement->nodeValue);
                        $oProperties->{$property}($oDateTime->getTimestamp());
                    } else {
                        $oProperties->{$property}($oElement->nodeValue);
                    }
                }
            }
        }
    }

    /**
     * Read Custom Properties
     * @param string $sPart
     */
    protected function loadCustomProperties($sPart)
    {
        $xmlReader = new XMLReader();
        $sPart = str_replace(' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"', '', $sPart);
        if ($xmlReader->getDomFromString($sPart)) {
            $pathMarkAsFinal = '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]/vt:bool';
            if (is_object($oElement = $xmlReader->getElement($pathMarkAsFinal))) {
                if ($oElement->nodeValue == 'true') {
                    $this->oPhpPresentation->getPresentationProperties()->markAsFinal(true);
                }
            }
        }
    }

    /**
     * Read View Properties
     * @param string $sPart
     */
    protected function loadViewProperties($sPart)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            $pathZoom = '/p:viewPr/p:slideViewPr/p:cSldViewPr/p:cViewPr/p:scale/a:sx';
            if (is_object($oElement = $xmlReader->getElement($pathZoom))) {
                if ($oElement->hasAttribute('d') && $oElement->hasAttribute('n')) {
                    $this->oPhpPresentation->getPresentationProperties()->setZoom($oElement->getAttribute('n') / $oElement->getAttribute('d'));
                }
            }
        }
    }

    /**
     * Extract all slides
     */
    protected function loadSlides($sPart)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            $fileRels = 'ppt/_rels/presentation.xml.rels';
            $this->loadRels($fileRels);
            // Load the Masterslides
            $this->loadMasterSlides($xmlReader, $fileRels);
            // Continue with loading the slides
            foreach ($xmlReader->getElements('/p:presentation/p:sldIdLst/p:sldId') as $oElement) {
                $rId = $oElement->getAttribute('r:id');
                $pathSlide = isset($this->arrayRels[$fileRels][$rId]) ? $this->arrayRels[$fileRels][$rId]['Target'] : '';
                if (!empty($pathSlide)) {
                    $pptSlide = $this->oZip->getFromName('ppt/' . $pathSlide);
                    if ($pptSlide !== false) {
                        $this->loadRels('ppt/slides/_rels/' . basename($pathSlide) . '.rels');
                        $this->loadSlide($pptSlide, basename($pathSlide));
                    }
                }
            }
        }
    }

    /**
     * Extract all MasterSlides
     * @param XMLReader $xmlReader
     * @param $fileRels
     */
    protected function loadMasterSlides(XMLReader $xmlReader, $fileRels)
    {
        // Get all the MasterSlide Id's from the presentation.xml file
        foreach ($xmlReader->getElements('/p:presentation/p:sldMasterIdLst/p:sldMasterId') as $oElement) {
            $rId = $oElement->getAttribute('r:id');
            // Get the path to the masterslide from the array with _rels files
            $pathMasterSlide = isset($this->arrayRels[$fileRels][$rId]) ?
                $this->arrayRels[$fileRels][$rId]['Target'] : '';
            if (!empty($pathMasterSlide)) {
                $pptMasterSlide = $this->oZip->getFromName('ppt/' . $pathMasterSlide);
                if ($pptMasterSlide !== false) {
                    $this->loadRels('ppt/slideMasters/_rels/' . basename($pathMasterSlide) . '.rels');
                    $this->loadMasterSlide($pptMasterSlide, basename($pathMasterSlide));
                }
            }
        }
    }

    /**
     * Extract data from slide
     * @param string $sPart
     * @param string $baseFile
     */
    protected function loadSlide($sPart, $baseFile)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            // Core
            $oSlide = $this->oPhpPresentation->createSlide();
            $this->oPhpPresentation->setActiveSlideIndex($this->oPhpPresentation->getSlideCount() - 1);
            $oSlide->setRelsIndex('ppt/slides/_rels/' . $baseFile . '.rels');

            // Background
            $oElement = $xmlReader->getElement('/p:sld/p:cSld/p:bg/p:bgPr');
            if ($oElement) {
                $oElementColor = $xmlReader->getElement('a:solidFill/a:srgbClr', $oElement);
                if ($oElementColor) {
                    // Color
                    $oColor = new Color();
                    $oColor->setRGB($oElementColor->hasAttribute('val') ? $oElementColor->getAttribute('val') : null);
                    // Background
                    $oBackground = new \PhpOffice\PhpPresentation\Slide\Background\Color();
                    $oBackground->setColor($oColor);
                    // Slide Background
                    $oSlide = $this->oPhpPresentation->getActiveSlide();
                    $oSlide->setBackground($oBackground);
                }
                $oElementImage = $xmlReader->getElement('a:blipFill/a:blip', $oElement);
                if ($oElementImage) {
                    $relImg = $this->arrayRels['ppt/slides/_rels/' . $baseFile . '.rels'][$oElementImage->getAttribute('r:embed')];
                    if (is_array($relImg)) {
                        // File
                        $pathImage = 'ppt/slides/' . $relImg['Target'];
                        $pathImage = explode('/', $pathImage);
                        foreach ($pathImage as $key => $partPath) {
                            if ($partPath == '..') {
                                unset($pathImage[$key - 1]);
                                unset($pathImage[$key]);
                            }
                        }
                        $pathImage = implode('/', $pathImage);
                        $contentImg = $this->oZip->getFromName($pathImage);

                        $tmpBkgImg = tempnam(sys_get_temp_dir(), 'PhpPresentationReaderPpt2007Bkg');
                        file_put_contents($tmpBkgImg, $contentImg);
                        // Background
                        $oBackground = new Image();
                        $oBackground->setPath($tmpBkgImg);
                        // Slide Background
                        $oSlide = $this->oPhpPresentation->getActiveSlide();
                        $oSlide->setBackground($oBackground);
                    }
                }
            }

            // Shapes
            foreach ($xmlReader->getElements('/p:sld/p:cSld/p:spTree/*') as $oNode) {
                switch ($oNode->tagName) {
                    case 'p:pic':
                        $this->loadShapeDrawing($xmlReader, $oNode, $oSlide);
                        break;
                    case 'p:sp':
                        $this->loadShapeRichText($xmlReader, $oNode, $oSlide);
                        break;
                    default:
                        //var_export($oNode->tagName);
                }
            }
            // Layout
            $oSlide = $this->oPhpPresentation->getActiveSlide();
            foreach ($this->arrayRels['ppt/slides/_rels/' . $baseFile . '.rels'] as $valueRel) {
                if ($valueRel['Type'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout') {
                    $layoutBasename = basename($valueRel['Target']);
                    if (array_key_exists($layoutBasename, $this->arraySlideLayouts)) {
                        $oSlide->setSlideLayout($this->arraySlideLayouts[$layoutBasename]);
                    }
                    break;
                }
            }
        }
    }

    private function loadMasterSlide($sPart, $baseFile)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            // Core
            $oSlideMaster = $this->oPhpPresentation->createMasterSlide();
            $oSlideMaster->setTextStyles(new TextStyle(false));
            $oSlideMaster->setRelsIndex('ppt/slideMasters/_rels/' . $baseFile . '.rels');

            // Background
            $oElement = $xmlReader->getElement('/p:sldMaster/p:cSld/p:bg');
            if ($oElement) {
                $this->loadSlideBackground($xmlReader, $oElement, $oSlideMaster);
            }

            // Shapes
            $arrayElements = $xmlReader->getElements('/p:sldMaster/p:cSld/p:spTree/*');
            if ($arrayElements) {
                $this->loadSlideShapes($oSlideMaster, $arrayElements, $xmlReader);
            }
            // Header & Footer

            // ColorMapping
            $colorMap = array();
            $oElement = $xmlReader->getElement('/p:sldMaster/p:clrMap');
            if ($oElement->hasAttributes()) {
                foreach ($oElement->attributes as $attr) {
                    $colorMap[$attr->nodeName] = $attr->nodeValue;
                }
                $oSlideMaster->colorMap->setMapping($colorMap);
            }

            // TextStyles
            $arrayElementTxStyles = $xmlReader->getElements('/p:sldMaster/p:txStyles/*');
            if ($arrayElementTxStyles) {
                foreach ($arrayElementTxStyles as $oElementTxStyle) {
                    $arrayElementsLvl = $xmlReader->getElements('/p:sldMaster/p:txStyles/'.$oElementTxStyle->nodeName.'/*');
                    foreach ($arrayElementsLvl as $oElementLvl) {
                        if ($oElementLvl->nodeName == 'a:extLst') {
                            continue;
                        }
                        $oRTParagraph = new Paragraph();

                        if ($oElementLvl->nodeName == 'a:defPPr') {
                            $level = 0;
                        } else {
                            $level = str_replace('a:lvl', '', $oElementLvl->nodeName);
                            $level = str_replace('pPr', '', $level);
                        }

                        if ($oElementLvl->hasAttribute('algn')) {
                            $oRTParagraph->getAlignment()->setHorizontal($oElementLvl->getAttribute('algn'));
                        }
                        if ($oElementLvl->hasAttribute('marL')) {
                            $val = $oElementLvl->getAttribute('marL');
                            $val = CommonDrawing::emuToPixels($val);
                            $oRTParagraph->getAlignment()->setMarginLeft($val);
                        }
                        if ($oElementLvl->hasAttribute('marR')) {
                            $val = $oElementLvl->getAttribute('marR');
                            $val = CommonDrawing::emuToPixels($val);
                            $oRTParagraph->getAlignment()->setMarginRight($val);
                        }
                        if ($oElementLvl->hasAttribute('indent')) {
                            $val = $oElementLvl->getAttribute('indent');
                            $val = CommonDrawing::emuToPixels($val);
                            $oRTParagraph->getAlignment()->setIndent($val);
                        }
                        $oElementLvlDefRPR = $xmlReader->getElement('a:defRPr', $oElementLvl);
                        if ($oElementLvlDefRPR) {
                            if ($oElementLvlDefRPR->hasAttribute('sz')) {
                                $oRTParagraph->getFont()->setSize($oElementLvlDefRPR->getAttribute('sz') / 100);
                            }
                            if ($oElementLvlDefRPR->hasAttribute('b') && $oElementLvlDefRPR->getAttribute('b') == 1) {
                                $oRTParagraph->getFont()->setBold(true);
                            }
                            if ($oElementLvlDefRPR->hasAttribute('i') && $oElementLvlDefRPR->getAttribute('i') == 1) {
                                $oRTParagraph->getFont()->setItalic(true);
                            }
                        }
                        $oElementSchemeColor = $xmlReader->getElement('a:defRPr/a:solidFill/a:schemeClr', $oElementLvl);
                        if ($oElementSchemeColor) {
                            if ($oElementSchemeColor->hasAttribute('val')) {
                                $oRTParagraph->getFont()->setColor(new SchemeColor())->getColor()->setValue($oElementSchemeColor->getAttribute('val'));
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
            }

            // Load the theme
            foreach ($this->arrayRels[$oSlideMaster->getRelsIndex()] as $arrayRel) {
                if ($arrayRel['Type'] == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
                    $pptTheme = $this->oZip->getFromName('ppt/' . substr($arrayRel['Target'], strrpos($arrayRel['Target'], '../') + 3));
                    if ($pptTheme !== false) {
                        $this->loadTheme($pptTheme, $oSlideMaster);
                    }
                    break;
                }
            }

            // Load the Layoutslide
            foreach ($xmlReader->getElements('/p:sldMaster/p:sldLayoutIdLst/p:sldLayoutId') as $oElement) {
                $rId = $oElement->getAttribute('r:id');
                // Get the path to the masterslide from the array with _rels files
                $pathLayoutSlide = isset($this->arrayRels[$oSlideMaster->getRelsIndex()][$rId]) ?
                    $this->arrayRels[$oSlideMaster->getRelsIndex()][$rId]['Target'] : '';
                if (!empty($pathLayoutSlide)) {
                    $pptLayoutSlide = $this->oZip->getFromName('ppt/' . substr($pathLayoutSlide, strrpos($pathLayoutSlide, '../') + 3));
                    if ($pptLayoutSlide !== false) {
                        $this->loadRels('ppt/slideLayouts/_rels/' . basename($pathLayoutSlide) . '.rels');
                        $oSlideMaster->addSlideLayout(
                            $this->loadLayoutSlide($pptLayoutSlide, basename($pathLayoutSlide), $oSlideMaster)
                        );
                    }
                }
            }
        }
    }

    private function loadLayoutSlide($sPart, $baseFile, SlideMaster $oSlideMaster)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            // Core
            $oSlideLayout = new SlideLayout($oSlideMaster);
            $oSlideLayout->setRelsIndex('ppt/slideLayouts/_rels/' . $baseFile . '.rels');

            // Name
            $oElement = $xmlReader->getElement('/p:sldLayout/p:cSld');
            if ($oElement && $oElement->hasAttribute('name')) {
                $oSlideLayout->setLayoutName($oElement->getAttribute('name'));
            }

            // Background
            $oElement = $xmlReader->getElement('/p:sldLayout/p:cSld/p:bg');
            if ($oElement) {
                $this->loadSlideBackground($xmlReader, $oElement, $oSlideLayout);
            }

            // ColorMapping
            $oElement = $xmlReader->getElement('/p:sldLayout/p:clrMapOvr/a:overrideClrMapping');
            if ($oElement && $oElement->hasAttributes()) {
                $colorMap = array();
                foreach ($oElement->attributes as $attr) {
                    $colorMap[$attr->nodeName] = $attr->nodeValue;
                }
                $oSlideLayout->colorMap->setMapping($colorMap);
            }

            // Shapes
            $oElements = $xmlReader->getElements('/p:sldLayout/p:cSld/p:spTree/*');
            if ($oElements) {
                $this->loadSlideShapes($oSlideLayout, $oElements, $xmlReader);
            }
            $this->arraySlideLayouts[$baseFile] = &$oSlideLayout;
            return $oSlideLayout;
        }
        return null;
    }

    /**
     * @param string $sPart
     * @param SlideMaster $oSlideMaster
     */
    private function loadTheme($sPart, SlideMaster $oSlideMaster)
    {
        $xmlReader = new XMLReader();
        if ($xmlReader->getDomFromString($sPart)) {
            $oElements = $xmlReader->getElements('/a:theme/a:themeElements/a:clrScheme/*');
            if ($oElements) {
                foreach ($oElements as $oElement) {
                    $oSchemeColor = new SchemeColor();
                    $oSchemeColor->setValue(str_replace('a:', '', $oElement->tagName));
                    $colorElement = $xmlReader->getElement('*', $oElement);
                    if ($colorElement) {
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

    /**
     * @param XMLReader $xmlReader
     * @param \DOMElement $oElement
     * @param AbstractSlide $oSlide
     */
    private function loadSlideBackground(XMLReader $xmlReader, \DOMElement $oElement, AbstractSlide $oSlide)
    {
        // Background color
        $oElementColor = $xmlReader->getElement('p:bgPr/a:solidFill/a:srgbClr', $oElement);
        if ($oElementColor) {
            // Color
            $oColor = new Color();
            $oColor->setRGB($oElementColor->hasAttribute('val') ? $oElementColor->getAttribute('val') : null);
            // Background
            $oBackground = new \PhpOffice\PhpPresentation\Slide\Background\Color();
            $oBackground->setColor($oColor);
            // Slide Background
            $oSlide->setBackground($oBackground);
        }

        // Background scheme color
        $oElementSchemeColor = $xmlReader->getElement('p:bgRef/a:schemeClr', $oElement);
        if ($oElementSchemeColor) {
            // Color
            $oColor = new SchemeColor();
            $oColor->setValue($oElementSchemeColor->hasAttribute('val') ? $oElementSchemeColor->getAttribute('val') : null);
            // Background
            $oBackground = new \PhpOffice\PhpPresentation\Slide\Background\SchemeColor();
            $oBackground->setSchemeColor($oColor);
            // Slide Background
            $oSlide->setBackground($oBackground);
        }

        // Background image
        $oElementImage = $xmlReader->getElement('p:bgPr/a:blipFill/a:blip', $oElement);
        if ($oElementImage) {
            $relImg = $this->arrayRels[$oSlide->getRelsIndex()][$oElementImage->getAttribute('r:embed')];
            if (is_array($relImg)) {
                // File
                $pathImage = 'ppt/slides/' . $relImg['Target'];
                $pathImage = explode('/', $pathImage);
                foreach ($pathImage as $key => $partPath) {
                    if ($partPath == '..') {
                        unset($pathImage[$key - 1]);
                        unset($pathImage[$key]);
                    }
                }
                $pathImage = implode('/', $pathImage);
                $contentImg = $this->oZip->getFromName($pathImage);

                $tmpBkgImg = tempnam(sys_get_temp_dir(), 'PhpPresentationReaderPpt2007Bkg');
                file_put_contents($tmpBkgImg, $contentImg);
                // Background
                $oBackground = new Image();
                $oBackground->setPath($tmpBkgImg);
                // Slide Background
                $oSlide->setBackground($oBackground);
            }
        }
    }

    /**
     *
     * @param XMLReader $document
     * @param \DOMElement $node
     * @param AbstractSlide $oSlide
     */
    protected function loadShapeDrawing(XMLReader $document, \DOMElement $node, AbstractSlide $oSlide)
    {
        // Core
        $oShape = new Gd();
        $oShape->getShadow()->setVisible(false);
        // Variables
        $fileRels = $oSlide->getRelsIndex();

        $oElement = $document->getElement('p:nvPicPr/p:cNvPr', $node);
        if ($oElement) {
            $oShape->setName($oElement->hasAttribute('name') ? $oElement->getAttribute('name') : '');
            $oShape->setDescription($oElement->hasAttribute('descr') ? $oElement->getAttribute('descr') : '');
        }

        $oElement = $document->getElement('p:blipFill/a:blip', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('r:embed') && isset($this->arrayRels[$fileRels][$oElement->getAttribute('r:embed')]['Target'])) {
                $pathImage = 'ppt/slides/' . $this->arrayRels[$fileRels][$oElement->getAttribute('r:embed')]['Target'];
                $pathImage = explode('/', $pathImage);
                foreach ($pathImage as $key => $partPath) {
                    if ($partPath == '..') {
                        unset($pathImage[$key - 1]);
                        unset($pathImage[$key]);
                    }
                }
                $pathImage = implode('/', $pathImage);
                $imageFile = $this->oZip->getFromName($pathImage);
                if (!empty($imageFile)) {
                    $oShape->setImageResource(imagecreatefromstring($imageFile));
                }
            }
        }

        $oElement = $document->getElement('p:spPr/a:xfrm', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('rot')) {
                $oShape->setRotation(CommonDrawing::angleToDegrees($oElement->getAttribute('rot')));
            }
        }

        $oElement = $document->getElement('p:spPr/a:xfrm/a:off', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('x')) {
                $oShape->setOffsetX(CommonDrawing::emuToPixels($oElement->getAttribute('x')));
            }
            if ($oElement->hasAttribute('y')) {
                $oShape->setOffsetY(CommonDrawing::emuToPixels($oElement->getAttribute('y')));
            }
        }

        $oElement = $document->getElement('p:spPr/a:xfrm/a:ext', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('cx')) {
                $oShape->setWidth(CommonDrawing::emuToPixels($oElement->getAttribute('cx')));
            }
            if ($oElement->hasAttribute('cy')) {
                $oShape->setHeight(CommonDrawing::emuToPixels($oElement->getAttribute('cy')));
            }
        }

        $oElement = $document->getElement('p:spPr/a:effectLst', $node);
        if ($oElement) {
            $oShape->getShadow()->setVisible(true);

            $oSubElement = $document->getElement('a:outerShdw', $oElement);
            if ($oSubElement) {
                if ($oSubElement->hasAttribute('blurRad')) {
                    $oShape->getShadow()->setBlurRadius(CommonDrawing::emuToPixels($oSubElement->getAttribute('blurRad')));
                }
                if ($oSubElement->hasAttribute('dist')) {
                    $oShape->getShadow()->setDistance(CommonDrawing::emuToPixels($oSubElement->getAttribute('dist')));
                }
                if ($oSubElement->hasAttribute('dir')) {
                    $oShape->getShadow()->setDirection(CommonDrawing::angleToDegrees($oSubElement->getAttribute('dir')));
                }
                if ($oSubElement->hasAttribute('algn')) {
                    $oShape->getShadow()->setAlignment($oSubElement->getAttribute('algn'));
                }
            }

            $oSubElement = $document->getElement('a:outerShdw/a:srgbClr', $oElement);
            if ($oSubElement) {
                if ($oSubElement->hasAttribute('val')) {
                    $oColor = new Color();
                    $oColor->setRGB($oSubElement->getAttribute('val'));
                    $oShape->getShadow()->setColor($oColor);
                }
            }

            $oSubElement = $document->getElement('a:outerShdw/a:srgbClr/a:alpha', $oElement);
            if ($oSubElement) {
                if ($oSubElement->hasAttribute('val')) {
                    $oShape->getShadow()->setAlpha((int)$oSubElement->getAttribute('val') / 1000);
                }
            }
        }

        $oSlide->addShape($oShape);
    }

    /**
     * @param XMLReader $document
     * @param \DOMElement $node
     * @param AbstractSlide $oSlide
     * @throws \Exception
     */
    protected function loadShapeRichText(XMLReader $document, \DOMElement $node, AbstractSlide $oSlide)
    {
        // Core
        $oShape = $oSlide->createRichTextShape();
        $oShape->setParagraphs(array());
        // Variables
        $fileRels = $oSlide->getRelsIndex();

        $oElement = $document->getElement('p:spPr/a:xfrm', $node);
        if ($oElement && $oElement->hasAttribute('rot')) {
            $oShape->setRotation(CommonDrawing::angleToDegrees($oElement->getAttribute('rot')));
        }

        $oElement = $document->getElement('p:spPr/a:xfrm/a:off', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('x')) {
                $oShape->setOffsetX(CommonDrawing::emuToPixels($oElement->getAttribute('x')));
            }
            if ($oElement->hasAttribute('y')) {
                $oShape->setOffsetY(CommonDrawing::emuToPixels($oElement->getAttribute('y')));
            }
        }

        $oElement = $document->getElement('p:spPr/a:xfrm/a:ext', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('cx')) {
                $oShape->setWidth(CommonDrawing::emuToPixels($oElement->getAttribute('cx')));
            }
            if ($oElement->hasAttribute('cy')) {
                $oShape->setHeight(CommonDrawing::emuToPixels($oElement->getAttribute('cy')));
            }
        }

        $oElement = $document->getElement('p:nvSpPr/p:nvPr/p:ph', $node);
        if ($oElement) {
            if ($oElement->hasAttribute('type')) {
                $placeholder = new Placeholder($oElement->getAttribute('type'));
                $oShape->setPlaceHolder($placeholder);
            }
        }

        $arrayElements = $document->getElements('p:txBody/a:p', $node);
        foreach ($arrayElements as $oElement) {
            // Core
            $oParagraph = $oShape->createParagraph();
            $oParagraph->setRichTextElements(array());

            $oSubElement = $document->getElement('a:pPr', $oElement);
            if ($oSubElement) {
                if ($oSubElement->hasAttribute('algn')) {
                    $oParagraph->getAlignment()->setHorizontal($oSubElement->getAttribute('algn'));
                }
                if ($oSubElement->hasAttribute('fontAlgn')) {
                    $oParagraph->getAlignment()->setVertical($oSubElement->getAttribute('fontAlgn'));
                }
                if ($oSubElement->hasAttribute('marL')) {
                    $oParagraph->getAlignment()->setMarginLeft(CommonDrawing::emuToPixels($oSubElement->getAttribute('marL')));
                }
                if ($oSubElement->hasAttribute('marR')) {
                    $oParagraph->getAlignment()->setMarginRight(CommonDrawing::emuToPixels($oSubElement->getAttribute('marR')));
                }
                if ($oSubElement->hasAttribute('indent')) {
                    $oParagraph->getAlignment()->setIndent(CommonDrawing::emuToPixels($oSubElement->getAttribute('indent')));
                }
                if ($oSubElement->hasAttribute('lvl')) {
                    $oParagraph->getAlignment()->setLevel($oSubElement->getAttribute('lvl'));
                }

                $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_NONE);

                $oElementBuFont = $document->getElement('a:buFont', $oSubElement);
                if ($oElementBuFont) {
                    if ($oElementBuFont->hasAttribute('typeface')) {
                        $oParagraph->getBulletStyle()->setBulletFont($oElementBuFont->getAttribute('typeface'));
                    }
                }
                $oElementBuChar = $document->getElement('a:buChar', $oSubElement);
                if ($oElementBuChar) {
                    $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
                    if ($oElementBuChar->hasAttribute('char')) {
                        $oParagraph->getBulletStyle()->setBulletChar($oElementBuChar->getAttribute('char'));
                    }
                }
                $oElementBuAutoNum = $document->getElement('a:buAutoNum', $oSubElement);
                if ($oElementBuAutoNum) {
                    $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_NUMERIC);
                    if ($oElementBuAutoNum->hasAttribute('type')) {
                        $oParagraph->getBulletStyle()->setBulletNumericStyle($oElementBuAutoNum->getAttribute('type'));
                    }
                    if ($oElementBuAutoNum->hasAttribute('startAt') && $oElementBuAutoNum->getAttribute('startAt') != 1) {
                        $oParagraph->getBulletStyle()->setBulletNumericStartAt($oElementBuAutoNum->getAttribute('startAt'));
                    }
                }
                $oElementBuClr = $document->getElement('a:buClr', $oSubElement);
                if ($oElementBuClr) {
                    $oColor = new Color();
                    /**
                     * @todo Create protected for reading Color
                     */
                    $oElementColor = $document->getElement('a:srgbClr', $oElementBuClr);
                    if ($oElementColor) {
                        $oColor->setRGB($oElementColor->hasAttribute('val') ? $oElementColor->getAttribute('val') : null);
                    }
                    $oParagraph->getBulletStyle()->setBulletColor($oColor);
                }
            }
            $arraySubElements = $document->getElements('(a:r|a:br)', $oElement);
            foreach ($arraySubElements as $oSubElement) {
                if ($oSubElement->tagName == 'a:br') {
                    $oParagraph->createBreak();
                }
                if ($oSubElement->tagName == 'a:r') {
                    $oElementrPr = $document->getElement('a:rPr', $oSubElement);
                    if (is_object($oElementrPr)) {
                        $oText = $oParagraph->createTextRun();

                        if ($oElementrPr->hasAttribute('b')) {
                            $oText->getFont()->setBold($oElementrPr->getAttribute('b') == 'true' ? true : false);
                        }
                        if ($oElementrPr->hasAttribute('i')) {
                            $oText->getFont()->setItalic($oElementrPr->getAttribute('i') == 'true' ? true : false);
                        }
                        if ($oElementrPr->hasAttribute('strike')) {
                            $oText->getFont()->setStrikethrough($oElementrPr->getAttribute('strike') == 'noStrike' ? false : true);
                        }
                        if ($oElementrPr->hasAttribute('sz')) {
                            $oText->getFont()->setSize((int)($oElementrPr->getAttribute('sz') / 100));
                        }
                        if ($oElementrPr->hasAttribute('u')) {
                            $oText->getFont()->setUnderline($oElementrPr->getAttribute('u'));
                        }
                        // Color
                        $oElementSrgbClr = $document->getElement('a:solidFill/a:srgbClr', $oElementrPr);
                        if (is_object($oElementSrgbClr) && $oElementSrgbClr->hasAttribute('val')) {
                            $oColor = new Color();
                            $oColor->setRGB($oElementSrgbClr->getAttribute('val'));
                            $oText->getFont()->setColor($oColor);
                        }
                        // Hyperlink
                        $oElementHlinkClick = $document->getElement('a:hlinkClick', $oElementrPr);
                        if (is_object($oElementHlinkClick)) {
                            if ($oElementHlinkClick->hasAttribute('tooltip')) {
                                $oText->getHyperlink()->setTooltip($oElementHlinkClick->getAttribute('tooltip'));
                            }
                            if ($oElementHlinkClick->hasAttribute('r:id') && isset($this->arrayRels[$fileRels][$oElementHlinkClick->getAttribute('r:id')]['Target'])) {
                                $oText->getHyperlink()->setUrl($this->arrayRels[$fileRels][$oElementHlinkClick->getAttribute('r:id')]['Target']);
                            }
                        }
                        //} else {
                        // $oText = $oParagraph->createText();

                        $oSubSubElement = $document->getElement('a:t', $oSubElement);
                        $oText->setText($oSubSubElement->nodeValue);
                    }
                }
            }
        }

        if (count($oShape->getParagraphs()) > 0) {
            $oShape->setActiveParagraph(0);
        }
    }

    /**
     *
     * @param string $fileRels
     * @return string
     */
    protected function loadRels($fileRels)
    {
        $sPart = $this->oZip->getFromName($fileRels);
        if ($sPart !== false) {
            $xmlReader = new XMLReader();
            if ($xmlReader->getDomFromString($sPart)) {
                foreach ($xmlReader->getElements('*') as $oNode) {
                    $this->arrayRels[$fileRels][$oNode->getAttribute('Id')] = array(
                        'Target' => $oNode->getAttribute('Target'),
                        'Type' => $oNode->getAttribute('Type'),
                    );
                }
            }
        }
    }

    /**
     * @param $oSlide
     * @param $oElements
     * @param $xmlReader
     * @internal param $baseFile
     */
    private function loadSlideShapes($oSlide, $oElements, $xmlReader)
    {
        foreach ($oElements as $oNode) {
            switch ($oNode->tagName) {
                case 'p:pic':
                    $this->loadShapeDrawing($xmlReader, $oNode, $oSlide);
                    break;
                case 'p:sp':
                    $this->loadShapeRichText($xmlReader, $oNode, $oSlide);
                    break;
                default:
                    //var_export($oNode->tagName);
            }
        }
    }
}
