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

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AbstractDrawing;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Drawing as ShapeDrawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide\Note;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpPresentation\Writer\ODPresentation;

/**
 * \PhpOffice\PhpPresentation\Writer\ODPresentation\Content
 */
class Content extends AbstractPart
{

    /**
     * Stores bullet styles for text shapes that include lists.
     *
     * @var array
     */
    private $arrStyleBullet    = array();

    /**
     * Stores paragraph information for text shapes.
     *
     * @var array
     */
    private $arrStyleParagraph = array();

    /**
     * Stores font styles for text shapes that include lists.
     *
     * @var array
     */
    private $arrStyleTextFont  = array();

    /**
     * Used to track the current shape ID.
     *
     * @var integer
     */
    private $shapeId;

    /**
     * Write content file to XML format
     *
     * @param  PhpPresentation $pPhpPresentation
     * @return string        XML Output
     * @throws \Exception
     */
    public function writePart(PhpPresentation $pPhpPresentation)
    {
        // Create XML writer
        $objWriter = $this->getXMLWriter();

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8');

        // office:document-content
        $objWriter->startElement('office:document-content');
        $objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
        $objWriter->writeAttribute('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0');
        $objWriter->writeAttribute('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
        $objWriter->writeAttribute('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
        $objWriter->writeAttribute('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
        $objWriter->writeAttribute('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0');
        $objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
        $objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $objWriter->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
        $objWriter->writeAttribute('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0');
        $objWriter->writeAttribute('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
        $objWriter->writeAttribute('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0');
        $objWriter->writeAttribute('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0');
        $objWriter->writeAttribute('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0');
        $objWriter->writeAttribute('xmlns:math', 'http://www.w3.org/1998/Math/MathML');
        $objWriter->writeAttribute('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0');
        $objWriter->writeAttribute('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0');
        $objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
        $objWriter->writeAttribute('xmlns:ooow', 'http://openoffice.org/2004/writer');
        $objWriter->writeAttribute('xmlns:oooc', 'http://openoffice.org/2004/calc');
        $objWriter->writeAttribute('xmlns:dom', 'http://www.w3.org/2001/xml-events');
        $objWriter->writeAttribute('xmlns:xforms', 'http://www.w3.org/2002/xforms');
        $objWriter->writeAttribute('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema');
        $objWriter->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
        $objWriter->writeAttribute('xmlns:smil', 'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0');
        $objWriter->writeAttribute('xmlns:anim', 'urn:oasis:names:tc:opendocument:xmlns:animation:1.0');
        $objWriter->writeAttribute('xmlns:rpt', 'http://openoffice.org/2005/report');
        $objWriter->writeAttribute('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2');
        $objWriter->writeAttribute('xmlns:rdfa', 'http://docs.oasis-open.org/opendocument/meta/rdfa#');
        $objWriter->writeAttribute('xmlns:field', 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0');
        $objWriter->writeAttribute('office:version', '1.2');

        // office:automatic-styles
        $objWriter->startElement('office:automatic-styles');

        $this->shapeId    = 0;
        $incSlide = 0;
        foreach ($pPhpPresentation->getAllSlides() as $pSlide) {
            // Slides
            $this->writeStyleSlide($objWriter, $pSlide, $incSlide);

            // Images
            $shapes = $pSlide->getShapeCollection();
            foreach ($shapes as $shape) {
                // Increment $this->shapeId
                ++$this->shapeId;

                // Check type
                if ($shape instanceof RichText) {
                    $this->writeTxtStyle($objWriter, $shape);
                }
                if ($shape instanceof AbstractDrawing) {
                    $this->writeDrawingStyle($objWriter, $shape);
                }
                if ($shape instanceof Line) {
                    $this->writeLineStyle($objWriter, $shape);
                }
                if ($shape instanceof Table) {
                    $this->writeTableStyle($objWriter, $shape);
                }
                if ($shape instanceof Group) {
                    $this->writeGroupStyle($objWriter, $shape);
                }
            }

            $incSlide++;
        }
        // Style : Bullet
        if (!empty($this->arrStyleBullet)) {
            foreach ($this->arrStyleBullet as $key => $item) {
                $oStyle   = $item['oStyle'];
                $arrLevel = explode(';', $item['level']);
                // style:style
                $objWriter->startElement('text:list-style');
                $objWriter->writeAttribute('style:name', 'L_' . $key);
                foreach ($arrLevel as $level) {
                    if ($level != '') {
                        $oAlign = $item['oAlign_' . $level];
                        // text:list-level-style-bullet
                        $objWriter->startElement('text:list-level-style-bullet');
                        $objWriter->writeAttribute('text:level', $level + 1);
                        $objWriter->writeAttribute('text:bullet-char', $oStyle->getBulletChar());
                        // style:list-level-properties
                        $objWriter->startElement('style:list-level-properties');
                        if ($oAlign->getIndent() < 0) {
                            $objWriter->writeAttribute('text:space-before', CommonDrawing::pixelsToCentimeters($oAlign->getMarginLeft() - (-1 * $oAlign->getIndent())) . 'cm');
                            $objWriter->writeAttribute('text:min-label-width', CommonDrawing::pixelsToCentimeters(-1 * $oAlign->getIndent()) . 'cm');
                        } else {
                            $objWriter->writeAttribute('text:space-before', (CommonDrawing::pixelsToCentimeters($oAlign->getMarginLeft() - $oAlign->getIndent())) . 'cm');
                            $objWriter->writeAttribute('text:min-label-width', CommonDrawing::pixelsToCentimeters($oAlign->getIndent()) . 'cm');
                        }

                        $objWriter->endElement();
                        // style:text-properties
                        $objWriter->startElement('style:text-properties');
                        $objWriter->writeAttribute('fo:font-family', $oStyle->getBulletFont());
                        $objWriter->writeAttribute('style:font-family-generic', 'swiss');
                        $objWriter->writeAttribute('style:use-window-font-color', 'true');
                        $objWriter->writeAttribute('fo:font-size', '100');
                        $objWriter->endElement();
                        $objWriter->endElement();
                    }
                }
                $objWriter->endElement();
            }
        }
        // Style : Paragraph
        if (!empty($this->arrStyleParagraph)) {
            foreach ($this->arrStyleParagraph as $key => $item) {
                // style:style
                $objWriter->startElement('style:style');
                $objWriter->writeAttribute('style:name', 'P_' . $key);
                $objWriter->writeAttribute('style:family', 'paragraph');
                // style:paragraph-properties
                $objWriter->startElement('style:paragraph-properties');
                switch ($item->getAlignment()->getHorizontal()) {
                    case Alignment::HORIZONTAL_LEFT:
                        $objWriter->writeAttribute('fo:text-align', 'left');
                        break;
                    case Alignment::HORIZONTAL_RIGHT:
                        $objWriter->writeAttribute('fo:text-align', 'right');
                        break;
                    case Alignment::HORIZONTAL_CENTER:
                        $objWriter->writeAttribute('fo:text-align', 'center');
                        break;
                    case Alignment::HORIZONTAL_JUSTIFY:
                        $objWriter->writeAttribute('fo:text-align', 'justify');
                        break;
                    case Alignment::HORIZONTAL_DISTRIBUTED:
                        $objWriter->writeAttribute('fo:text-align', 'justify');
                        break;
                    default:
                        $objWriter->writeAttribute('fo:text-align', 'left');
                        break;
                }
                $objWriter->endElement();
                $objWriter->endElement();
            }
        }
        // Style : Text : Font
        if (!empty($this->arrStyleTextFont)) {
            foreach ($this->arrStyleTextFont as $key => $item) {
                // style:style
                $objWriter->startElement('style:style');
                $objWriter->writeAttribute('style:name', 'T_' . $key);
                $objWriter->writeAttribute('style:family', 'text');
                // style:text-properties
                $objWriter->startElement('style:text-properties');
                $objWriter->writeAttribute('fo:color', '#' . $item->getColor()->getRGB());
                $objWriter->writeAttribute('fo:font-family', $item->getName());
                $objWriter->writeAttribute('fo:font-size', $item->getSize() . 'pt');
                // @todo : fo:font-style
                if ($item->isBold()) {
                    $objWriter->writeAttribute('fo:font-weight', 'bold');
                }
                // @todo : style:text-underline-style
                $objWriter->endElement();
                $objWriter->endElement();
            }
        }
        $objWriter->endElement();

        //===============================================
        // Body
        //===============================================
        // office:body
        $objWriter->startElement('office:body');
        // office:presentation
        $objWriter->startElement('office:presentation');

        // Write slides
        $slideCount = $pPhpPresentation->getSlideCount();
        $this->shapeId    = 0;
        for ($i = 0; $i < $slideCount; ++$i) {
            $pSlide = $pPhpPresentation->getSlide($i);
            $objWriter->startElement('draw:page');
            $objWriter->writeAttribute('draw:name', 'page' . $i);
            $objWriter->writeAttribute('draw:master-page-name', 'Standard');
            $objWriter->writeAttribute('draw:style-name', 'stylePage' . $i);
            // Images
            $shapes = $pSlide->getShapeCollection();
            foreach ($shapes as $shape) {
                // Increment $this->shapeId
                ++$this->shapeId;

                // Check type
                if ($shape instanceof RichText) {
                    $this->writeShapeTxt($objWriter, $shape);
                } elseif ($shape instanceof Table) {
                    $this->writeShapeTable($objWriter, $shape);
                } elseif ($shape instanceof Line) {
                    $this->writeShapeLine($objWriter, $shape);
                } elseif ($shape instanceof Chart) {
                    $this->writeShapeChart($objWriter, $shape);
                } elseif ($shape instanceof AbstractDrawing) {
                    $this->writeShapePic($objWriter, $shape);
                } elseif ($shape instanceof Group) {
                    $this->writeShapeGroup($objWriter, $shape);
                }
            }
            // Slide Note
            if ($pSlide->getNote() instanceof Note) {
                $this->writeSlideNote($objWriter, $pSlide->getNote());
            }
            
            $objWriter->endElement();
        }
        $objWriter->endElement();
        $objWriter->endElement();
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write picture
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter
     * @param \PhpOffice\PhpPresentation\Shape\AbstractDrawing $shape
     */
    public function writeShapePic(XMLWriter $objWriter, AbstractDrawing $shape)
    {
        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('draw:name', $shape->getName());
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getWidth()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getHeight()), 3) . 'cm');
        $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('draw:style-name', 'gr' . $this->shapeId);
        // draw:image
        $objWriter->startElement('draw:image');
        if ($shape instanceof ShapeDrawing) {
            $objWriter->writeAttribute('xlink:href', 'Pictures/' . md5($shape->getPath()) . '.' . $shape->getExtension());
        } elseif ($shape instanceof MemoryDrawing) {
            $objWriter->writeAttribute('xlink:href', 'Pictures/' . $shape->getIndexedFilename());
        }
        $objWriter->writeAttribute('xlink:type', 'simple');
        $objWriter->writeAttribute('xlink:show', 'embed');
        $objWriter->writeAttribute('xlink:actuate', 'onLoad');
        $objWriter->writeElement('text:p');
        $objWriter->endElement();
        
        if ($shape->hasHyperlink()) {
            // office:event-listeners
            $objWriter->startElement('office:event-listeners');
            // presentation:event-listener
            $objWriter->startElement('presentation:event-listener');
            $objWriter->writeAttribute('script:event-name', 'dom:click');
            $objWriter->writeAttribute('presentation:action', 'show');
            $objWriter->writeAttribute('xlink:href', $shape->getHyperlink()->getUrl());
            $objWriter->writeAttribute('xlink:type', 'simple');
            $objWriter->writeAttribute('xlink:show', 'embed');
            $objWriter->writeAttribute('xlink:actuate', 'onRequest');
            // > presentation:event-listener
            $objWriter->endElement();
            // > office:event-listeners
            $objWriter->endElement();
        }
        
        $objWriter->endElement();
    }

    /**
     * Write text
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter
     * @param \PhpOffice\PhpPresentation\Shape\RichText $shape
     */
    public function writeShapeTxt(XMLWriter $objWriter, RichText $shape)
    {
        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('draw:style-name', 'gr' . $this->shapeId);
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getWidth()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getHeight()), 3) . 'cm');
        $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetY()), 3) . 'cm');
        // draw:text-box
        $objWriter->startElement('draw:text-box');
        
        $paragraphs             = $shape->getParagraphs();
        $paragraphId            = 0;
        $sCstShpLastBullet      = '';
        $iCstShpLastBulletLvl   = 0;
        $bCstShpHasBullet       = false;

        foreach ($paragraphs as $paragraph) {
            // Close the bullet list
            if ($sCstShpLastBullet != 'bullet' && $bCstShpHasBullet === true) {
                for ($iInc = $iCstShpLastBulletLvl; $iInc >= 0; $iInc--) {
                    // text:list-item
                    $objWriter->endElement();
                    // text:list
                    $objWriter->endElement();
                }
            }
            //===============================================
            // Paragraph
            //===============================================
            if ($paragraph->getBulletStyle()->getBulletType() == 'none') {
                ++$paragraphId;
                // text:p
                $objWriter->startElement('text:p');
                $objWriter->writeAttribute('text:style-name', 'P_' . $paragraph->getHashCode());

                // Loop trough rich text elements
                $richtexts  = $paragraph->getRichTextElements();
                $richtextId = 0;
                foreach ($richtexts as $richtext) {
                    ++$richtextId;
                    if ($richtext instanceof TextElement || $richtext instanceof Run) {
                        // text:span
                        $objWriter->startElement('text:span');
                        if ($richtext instanceof Run) {
                            $objWriter->writeAttribute('text:style-name', 'T_' . $richtext->getFont()->getHashCode());
                        }
                        if ($richtext->hasHyperlink() === true && $richtext->getHyperlink()->getUrl() != '') {
                            // text:a
                            $objWriter->startElement('text:a');
                            $objWriter->writeAttribute('xlink:href', $richtext->getHyperlink()->getUrl());
                            $objWriter->text($richtext->getText());
                            $objWriter->endElement();
                        } else {
                            $objWriter->text($richtext->getText());
                        }
                        $objWriter->endElement();
                    } elseif ($richtext instanceof BreakElement) {
                        // text:span
                        $objWriter->startElement('text:span');
                        // text:line-break
                        $objWriter->startElement('text:line-break');
                        $objWriter->endElement();
                        $objWriter->endElement();
                    } else {
                        //echo '<pre>'.print_r($richtext, true).'</pre>';
                    }
                }
                $objWriter->endElement();
            //===============================================
            // Bullet list
            //===============================================
            } elseif ($paragraph->getBulletStyle()->getBulletType() == 'bullet') {
                $bCstShpHasBullet = true;
                // Open the bullet list
                if ($sCstShpLastBullet != 'bullet' || ($sCstShpLastBullet == $paragraph->getBulletStyle()->getBulletType() && $iCstShpLastBulletLvl < $paragraph->getAlignment()->getLevel())) {
                    // text:list
                    $objWriter->startElement('text:list');
                    $objWriter->writeAttribute('text:style-name', 'L_' . $paragraph->getBulletStyle()->getHashCode());
                }
                if ($sCstShpLastBullet == 'bullet') {
                    if ($iCstShpLastBulletLvl == $paragraph->getAlignment()->getLevel()) {
                        // text:list-item
                        $objWriter->endElement();
                    } elseif ($iCstShpLastBulletLvl > $paragraph->getAlignment()->getLevel()) {
                        // text:list-item
                        $objWriter->endElement();
                        // text:list
                        $objWriter->endElement();
                        // text:list-item
                        $objWriter->endElement();
                    }
                }

                // text:list-item
                $objWriter->startElement('text:list-item');
                ++$paragraphId;
                // text:p
                $objWriter->startElement('text:p');
                $objWriter->writeAttribute('text:style-name', 'P_' . $paragraph->getHashCode());

                // Loop trough rich text elements
                $richtexts  = $paragraph->getRichTextElements();
                $richtextId = 0;
                foreach ($richtexts as $richtext) {
                    ++$richtextId;
                    if ($richtext instanceof TextElement || $richtext instanceof Run) {
                        // text:span
                        $objWriter->startElement('text:span');
                        if ($richtext instanceof Run) {
                            $objWriter->writeAttribute('text:style-name', 'T_' . $richtext->getFont()->getHashCode());
                        }
                        if ($richtext->hasHyperlink() === true && $richtext->getHyperlink()->getUrl() != '') {
                            // text:a
                            $objWriter->startElement('text:a');
                            $objWriter->writeAttribute('xlink:href', $richtext->getHyperlink()->getUrl());
                            $objWriter->text($richtext->getText());
                            $objWriter->endElement();
                        } else {
                            $objWriter->text($richtext->getText());
                        }
                        $objWriter->endElement();
                    } elseif ($richtext instanceof BreakElement) {
                        // text:span
                        $objWriter->startElement('text:span');
                        // text:line-break
                        $objWriter->startElement('text:line-break');
                        $objWriter->endElement();
                        $objWriter->endElement();
                    } else {
                        //echo '<pre>'.print_r($richtext, true).'</pre>';
                    }
                }
                $objWriter->endElement();
            }
            $sCstShpLastBullet      = $paragraph->getBulletStyle()->getBulletType();
            $iCstShpLastBulletLvl = $paragraph->getAlignment()->getLevel();
        }

        // Close the bullet list
        if ($sCstShpLastBullet == 'bullet' && $bCstShpHasBullet === true) {
            for ($iInc = $iCstShpLastBulletLvl; $iInc >= 0; $iInc--) {
                // text:list-item
                $objWriter->endElement();
                // text:list
                $objWriter->endElement();
            }
        }
        
        // > draw:text-box
        $objWriter->endElement();
        // > draw:frame
        $objWriter->endElement();
    }

    /**
     * @param XMLWriter $objWriter
     * @param Line $shape
     */
    public function writeShapeLine(XMLWriter $objWriter, Line $shape)
    {
        // draw:line
        $objWriter->startElement('draw:line');
        $objWriter->writeAttribute('draw:style-name', 'gr' . $this->shapeId);
        $objWriter->writeAttribute('svg:x1', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y1', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('svg:x2', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetX()+$shape->getWidth()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y2', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetY()+$shape->getHeight()), 3) . 'cm');

            // text:p
            $objWriter->writeElement('text:p');

        $objWriter->endElement();
    }

    /**
     * Write table Shape
     * @param XMLWriter $objWriter
     * @param Table $shape
     */
    public function writeShapeTable(XMLWriter $objWriter, Table $shape)
    {
        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getHeight()), 3) . 'cm');
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getWidth()), 3) . 'cm');
        
        // table:table
        $objWriter->startElement('table:table');
        
        foreach ($shape->getRows() as $keyRow => $shapeRow) {
            // table:table-row
            $objWriter->startElement('table:table-row');
            $objWriter->writeAttribute('table:style-name', 'gr'.$this->shapeId.'r'.$keyRow);
            //@todo getFill
            
            $numColspan = 0;
            foreach ($shapeRow->getCells() as $keyCell => $shapeCell) {
                if ($numColspan == 0) {
                    // table:table-cell
                    $objWriter->startElement('table:table-cell');
                    $objWriter->writeAttribute('table:style-name', 'gr' . $this->shapeId.'r'.$keyRow.'c'.$keyCell);
                    if ($shapeCell->getColspan() > 1) {
                        $objWriter->writeAttribute('table:number-columns-spanned', $shapeCell->getColspan());
                        $numColspan = $shapeCell->getColspan() - 1;
                    }
                    
                    // text:p
                    $objWriter->startElement('text:p');
                    
                    // text:span
                    foreach ($shapeCell->getParagraphs() as $shapeParagraph) {
                        foreach ($shapeParagraph->getRichTextElements() as $shapeRichText) {
                            if ($shapeRichText instanceof TextElement || $shapeRichText instanceof Run) {
                                // text:span
                                $objWriter->startElement('text:span');
                                if ($shapeRichText instanceof Run) {
                                    $objWriter->writeAttribute('text:style-name', 'T_' . $shapeRichText->getFont()->getHashCode());
                                }
                                if ($shapeRichText->hasHyperlink() === true && $shapeRichText->getHyperlink()->getUrl() != '') {
                                    // text:a
                                    $objWriter->startElement('text:a');
                                    $objWriter->writeAttribute('xlink:href', $shapeRichText->getHyperlink()->getUrl());
                                    $objWriter->text($shapeRichText->getText());
                                    $objWriter->endElement();
                                } else {
                                    $objWriter->text($shapeRichText->getText());
                                }
                                $objWriter->endElement();
                            } elseif ($shapeRichText instanceof BreakElement) {
                                // text:span
                                $objWriter->startElement('text:span');
                                // text:line-break
                                $objWriter->startElement('text:line-break');
                                $objWriter->endElement();
                                $objWriter->endElement();
                            }
                        }
                    }
                    
                    // > text:p
                    $objWriter->endElement();
                    
                    // > table:table-cell
                    $objWriter->endElement();
                } else {
                    // table:covered-table-cell
                    $objWriter->writeElement('table:covered-table-cell');
                    $numColspan--;
                }
            }
            // > table:table-row
            $objWriter->endElement();
        }
        // > table:table
        $objWriter->endElement();
        // > draw:frame
        $objWriter->endElement();
    }
    
    /**
     * Write table Chart
     * @param XMLWriter $objWriter
     * @param Chart $shape
     */
    public function writeShapeChart(XMLWriter $objWriter, Chart $shape)
    {
        $parentWriter = $this->getParentWriter();
        if (!$parentWriter instanceof ODPresentation) {
            throw new \Exception('The $parentWriter is not an instance of \PhpOffice\PhpPresentation\Writer\ODPresentation');
        }
        $parentWriter->chartArray[$this->shapeId] = $shape;
        
        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('draw:name', $shape->getTitle()->getText());
        $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getHeight()), 3) . 'cm');
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters($shape->getWidth()), 3) . 'cm');
    
        // draw:object
        $objWriter->startElement('draw:object');
        $objWriter->writeAttribute('xlink:href', './Object '.$this->shapeId);
        $objWriter->writeAttribute('xlink:type', 'simple');
        $objWriter->writeAttribute('xlink:show', 'embed');
        
        // > draw:object
        $objWriter->endElement();
        // > draw:frame
        $objWriter->endElement();
    }

    /**
     * Writes a group of shapes
     *
     * @param XMLWriter $objWriter
     * @param Group $group
     */
    public function writeShapeGroup(XMLWriter $objWriter, Group $group)
    {
        // draw:g
        $objWriter->startElement('draw:g');

        $shapes = $group->getShapeCollection();
        foreach ($shapes as $shape) {
            // Increment $this->shapeId
            ++$this->shapeId;

            // Check type
            if ($shape instanceof RichText) {
                $this->writeShapeTxt($objWriter, $shape);
            } elseif ($shape instanceof Table) {
                $this->writeShapeTable($objWriter, $shape);
            } elseif ($shape instanceof Line) {
                $this->writeShapeLine($objWriter, $shape);
            } elseif ($shape instanceof Chart) {
                $this->writeShapeChart($objWriter, $shape);
            } elseif ($shape instanceof AbstractDrawing) {
                $this->writeShapePic($objWriter, $shape);
            } elseif ($shape instanceof Group) {
                $this->writeShapeGroup($objWriter, $shape);
            }
        }

        $objWriter->endElement(); // draw:g
    }

    /**
     * Writes the style information for a group of shapes
     *
     * @param XMLWriter $objWriter
     * @param Group $group
     */
    public function writeGroupStyle(XMLWriter $objWriter, Group $group)
    {
        $shapes = $group->getShapeCollection();
        foreach ($shapes as $shape) {
            // Increment $this->shapeId
            ++$this->shapeId;

            // Check type
            if ($shape instanceof RichText) {
                $this->writeTxtStyle($objWriter, $shape);
            }
            if ($shape instanceof AbstractDrawing) {
                $this->writeDrawingStyle($objWriter, $shape);
            }
            if ($shape instanceof Line) {
                $this->writeLineStyle($objWriter, $shape);
            }
            if ($shape instanceof Table) {
                $this->writeTableStyle($objWriter, $shape);
            }
        }
    }

    /**
     * Write the default style information for a RichText shape
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter
     * @param \PhpOffice\PhpPresentation\Shape\RichText $shape
     */
    public function writeTxtStyle(XMLWriter $objWriter, RichText $shape)
    {
        // style:style
        $objWriter->startElement('style:style');
        $objWriter->writeAttribute('style:name', 'gr' . $this->shapeId);
        $objWriter->writeAttribute('style:family', 'graphic');
        $objWriter->writeAttribute('style:parent-style-name', 'standard');
        // style:graphic-properties
        $objWriter->startElement('style:graphic-properties');
        if ($shape->getShadow()->isVisible()) {
            $this->writeStylePartShadow($objWriter, $shape->getShadow());
        }
        if (is_bool($shape->hasAutoShrinkVertical())) {
            $objWriter->writeAttribute('draw:auto-grow-height', var_export($shape->hasAutoShrinkVertical(), true));
        }
        if (is_bool($shape->hasAutoShrinkHorizontal())) {
            $objWriter->writeAttribute('draw:auto-grow-width', var_export($shape->hasAutoShrinkHorizontal(), true));
        }
        // Fill
        switch ($shape->getFill()->getFillType()) {
            case Fill::FILL_GRADIENT_LINEAR:
            case Fill::FILL_GRADIENT_PATH:
                $objWriter->writeAttribute('draw:fill', 'gradient');
                $objWriter->writeAttribute('draw:fill-gradient-name', 'gradient_'.$shape->getFill()->getHashCode());
                break;
            case Fill::FILL_SOLID:
                $objWriter->writeAttribute('draw:fill', 'solid');
                $objWriter->writeAttribute('draw:fill-color', '#'.$shape->getFill()->getStartColor()->getRGB());
                break;
            case Fill::FILL_NONE:
            default:
                $objWriter->writeAttribute('draw:fill', 'none');
                $objWriter->writeAttribute('draw:fill-color', '#'.$shape->getFill()->getStartColor()->getRGB());
                break;
        }
        // Border
        if ($shape->getBorder()->getLineStyle() == Border::LINE_NONE) {
            $objWriter->writeAttribute('draw:stroke', 'none');
        } else {
            $objWriter->writeAttribute('svg:stroke-color', '#'.$shape->getBorder()->getColor()->getRGB());
            $objWriter->writeAttribute('svg:stroke-width', number_format(CommonDrawing::pointsToCentimeters($shape->getBorder()->getLineWidth()), 3, '.', '').'cm');
            switch ($shape->getBorder()->getDashStyle()) {
                case Border::DASH_SOLID:
                    $objWriter->writeAttribute('draw:stroke', 'solid');
                    break;
                case Border::DASH_DASH:
                case Border::DASH_DASHDOT:
                case Border::DASH_DOT:
                case Border::DASH_LARGEDASH:
                case Border::DASH_LARGEDASHDOT:
                case Border::DASH_LARGEDASHDOTDOT:
                case Border::DASH_SYSDASH:
                case Border::DASH_SYSDASHDOT:
                case Border::DASH_SYSDASHDOTDOT:
                case Border::DASH_SYSDOT:
                    $objWriter->writeAttribute('draw:stroke', 'dash');
                    $objWriter->writeAttribute('draw:stroke-dash', 'strokeDash_'.$shape->getBorder()->getDashStyle());
                    break;
                default:
                    $objWriter->writeAttribute('draw:stroke', 'none');
                    break;
            }
        }

        $objWriter->writeAttribute('fo:wrap-option', 'wrap');
        // > style:graphic-properties
        $objWriter->endElement();
        // > style:style
        $objWriter->endElement();

        $paragraphs  = $shape->getParagraphs();
        $paragraphId = 0;
        foreach ($paragraphs as $paragraph) {
            ++$paragraphId;

            // Style des paragraphes
            if (!isset($this->arrStyleParagraph[$paragraph->getHashCode()])) {
                $this->arrStyleParagraph[$paragraph->getHashCode()] = $paragraph;
            }

            // Style des listes
            if (!isset($this->arrStyleBullet[$paragraph->getBulletStyle()->getHashCode()])) {
                $this->arrStyleBullet[$paragraph->getBulletStyle()->getHashCode()]['oStyle'] = $paragraph->getBulletStyle();
                $this->arrStyleBullet[$paragraph->getBulletStyle()->getHashCode()]['level']  = '';
            }
            if (strpos($this->arrStyleBullet[$paragraph->getBulletStyle()->getHashCode()]['level'], ';' . $paragraph->getAlignment()->getLevel()) === false) {
                $this->arrStyleBullet[$paragraph->getBulletStyle()->getHashCode()]['level'] .= ';' . $paragraph->getAlignment()->getLevel();
                $this->arrStyleBullet[$paragraph->getBulletStyle()->getHashCode()]['oAlign_' . $paragraph->getAlignment()->getLevel()] = $paragraph->getAlignment();
            }

            $richtexts  = $paragraph->getRichTextElements();
            $richtextId = 0;
            foreach ($richtexts as $richtext) {
                ++$richtextId;
                // Not a line break
                if ($richtext instanceof Run) {
                    // Style des font text
                    if (!isset($this->arrStyleTextFont[$richtext->getFont()->getHashCode()])) {
                        $this->arrStyleTextFont[$richtext->getFont()->getHashCode()] = $richtext->getFont();
                    }
                }
            }
        }
    }

    /**
     * Write the default style information for an AbstractDrawing
     *
     * @param \PhpOffice\Common\XMLWriter $objWriter
     * @param \PhpOffice\PhpPresentation\Shape\AbstractDrawing $shape
     */
    public function writeDrawingStyle(XMLWriter $objWriter, AbstractDrawing $shape)
    {
        // style:style
        $objWriter->startElement('style:style');
        $objWriter->writeAttribute('style:name', 'gr' . $this->shapeId);
        $objWriter->writeAttribute('style:family', 'graphic');
        $objWriter->writeAttribute('style:parent-style-name', 'standard');

        // style:graphic-properties
        $objWriter->startElement('style:graphic-properties');
        $objWriter->writeAttribute('draw:stroke', 'none');
        $objWriter->writeAttribute('draw:fill', 'none');
        if ($shape->getShadow()->isVisible()) {
            $this->writeStylePartShadow($objWriter, $shape->getShadow());
        }
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write the default style information for a Line shape.
     *
     * @param XMLWriter $objWriter
     * @param Line $shape
     */
    public function writeLineStyle(XMLWriter $objWriter, Line $shape)
    {
        // style:style
        $objWriter->startElement('style:style');
        $objWriter->writeAttribute('style:name', 'gr' . $this->shapeId);
        $objWriter->writeAttribute('style:family', 'graphic');
        $objWriter->writeAttribute('style:parent-style-name', 'standard');

        // style:graphic-properties
        $objWriter->startElement('style:graphic-properties');
        $objWriter->writeAttribute('draw:fill', 'none');
        switch ($shape->getBorder()->getLineStyle()) {
            case Border::LINE_NONE:
                $objWriter->writeAttribute('draw:stroke', 'none');
                break;
            case Border::LINE_SINGLE:
                $objWriter->writeAttribute('draw:stroke', 'solid');
                break;
            default:
                $objWriter->writeAttribute('draw:stroke', 'none');
                break;
        }
        $objWriter->writeAttribute('svg:stroke-color', '#'.$shape->getBorder()->getColor()->getRGB());
        $objWriter->writeAttribute('svg:stroke-width', Text::numberFormat(CommonDrawing::pixelsToCentimeters((CommonDrawing::pointsToPixels($shape->getBorder()->getLineWidth()))), 3).'cm');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write the default style information for a Table shape
     *
     * @param XMLWriter $objWriter
     * @param Table $shape
     */
    public function writeTableStyle(XMLWriter $objWriter, Table $shape)
    {
        foreach ($shape->getRows() as $keyRow => $shapeRow) {
            // style:style
            $objWriter->startElement('style:style');
            $objWriter->writeAttribute('style:name', 'gr' . $this->shapeId.'r'.$keyRow);
            $objWriter->writeAttribute('style:family', 'table-row');

            // style:table-row-properties
            $objWriter->startElement('style:table-row-properties');
            $objWriter->writeAttribute('style:row-height', Text::numberFormat(CommonDrawing::pixelsToCentimeters(CommonDrawing::pointsToPixels($shapeRow->getHeight())), 3).'cm');
            $objWriter->endElement();

            $objWriter->endElement();

            foreach ($shapeRow->getCells() as $keyCell => $shapeCell) {
                // style:style
                $objWriter->startElement('style:style');
                $objWriter->writeAttribute('style:name', 'gr' . $this->shapeId.'r'.$keyRow.'c'.$keyCell);
                $objWriter->writeAttribute('style:family', 'table-cell');

                // style:graphic-properties
                $objWriter->startElement('style:graphic-properties');
                if ($shapeCell->getFill()->getFillType() == Fill::FILL_SOLID) {
                    $objWriter->writeAttribute('draw:fill', 'solid');
                    $objWriter->writeAttribute('draw:fill-color', '#'.$shapeCell->getFill()->getStartColor()->getRGB());
                }
                if ($shapeCell->getFill()->getFillType() == Fill::FILL_GRADIENT_LINEAR) {
                    $objWriter->writeAttribute('draw:fill', 'gradient');
                    $objWriter->writeAttribute('draw:fill-gradient-name', 'gradient_'.$shapeCell->getFill()->getHashCode());
                }
                $objWriter->endElement();
                // <style:graphic-properties

                // style:paragraph-properties
                $objWriter->startElement('style:paragraph-properties');
                if ($shapeCell->getBorders()->getBottom()->getHashCode() == $shapeCell->getBorders()->getTop()->getHashCode()
                    && $shapeCell->getBorders()->getBottom()->getHashCode() == $shapeCell->getBorders()->getLeft()->getHashCode()
                    && $shapeCell->getBorders()->getBottom()->getHashCode() == $shapeCell->getBorders()->getRight()->getHashCode()) {
                    $lineStyle = 'none';
                    $lineWidth = Text::numberFormat($shapeCell->getBorders()->getBottom()->getLineWidth() / 1.75, 2);
                    $lineColor = $shapeCell->getBorders()->getBottom()->getColor()->getRGB();
                    switch ($shapeCell->getBorders()->getBottom()->getLineStyle()) {
                        case Border::LINE_SINGLE:
                            $lineStyle = 'solid';
                    }
                    $objWriter->writeAttribute('fo:border', $lineWidth.'pt '.$lineStyle.' #'.$lineColor);
                } else {
                    $lineStyle = 'none';
                    $lineWidth = Text::numberFormat($shapeCell->getBorders()->getBottom()->getLineWidth() / 1.75, 2);
                    $lineColor = $shapeCell->getBorders()->getBottom()->getColor()->getRGB();
                    switch ($shapeCell->getBorders()->getBottom()->getLineStyle()) {
                        case Border::LINE_SINGLE:
                            $lineStyle = 'solid';
                    }
                    $objWriter->writeAttribute('fo:border-bottom', $lineWidth.'pt '.$lineStyle.' #'.$lineColor);
                    // TOP
                    $lineStyle = 'none';
                    $lineWidth = Text::numberFormat($shapeCell->getBorders()->getTop()->getLineWidth() / 1.75, 2);
                    $lineColor = $shapeCell->getBorders()->getTop()->getColor()->getRGB();
                    switch ($shapeCell->getBorders()->getTop()->getLineStyle()) {
                        case Border::LINE_SINGLE:
                            $lineStyle = 'solid';
                    }
                    $objWriter->writeAttribute('fo:border-top', $lineWidth.'pt '.$lineStyle.' #'.$lineColor);
                    // RIGHT
                    $lineStyle = 'none';
                    $lineWidth = Text::numberFormat($shapeCell->getBorders()->getRight()->getLineWidth() / 1.75, 2);
                    $lineColor = $shapeCell->getBorders()->getRight()->getColor()->getRGB();
                    switch ($shapeCell->getBorders()->getRight()->getLineStyle()) {
                        case Border::LINE_SINGLE:
                            $lineStyle = 'solid';
                    }
                    $objWriter->writeAttribute('fo:border-right', $lineWidth.'pt '.$lineStyle.' #'.$lineColor);
                    // LEFT
                    $lineStyle = 'none';
                    $lineWidth = Text::numberFormat($shapeCell->getBorders()->getLeft()->getLineWidth() / 1.75, 2);
                    $lineColor = $shapeCell->getBorders()->getLeft()->getColor()->getRGB();
                    switch ($shapeCell->getBorders()->getLeft()->getLineStyle()) {
                        case Border::LINE_SINGLE:
                            $lineStyle = 'solid';
                    }
                    $objWriter->writeAttribute('fo:border-left', $lineWidth.'pt '.$lineStyle.' #'.$lineColor);
                }
                $objWriter->endElement();

                $objWriter->endElement();

                foreach ($shapeCell->getParagraphs() as $shapeParagraph) {
                    foreach ($shapeParagraph->getRichTextElements() as $shapeRichText) {
                        if ($shapeRichText instanceof Run) {
                            // Style des font text
                            if (!isset($this->arrStyleTextFont[$shapeRichText->getFont()->getHashCode()])) {
                                $this->arrStyleTextFont[$shapeRichText->getFont()->getHashCode()] = $shapeRichText->getFont();
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Write the slide note
     * @param XMLWriter $objWriter
     * @param \PhpOffice\PhpPresentation\Slide\Note $note
     */
    public function writeSlideNote(XMLWriter $objWriter, Note $note)
    {
        $shapesNote = $note->getShapeCollection();
        if (count($shapesNote) > 0) {
            $objWriter->startElement('presentation:notes');
            
            foreach ($shapesNote as $shape) {
                // Increment $this->shapeId
                ++$this->shapeId;
                
                if ($shape instanceof RichText) {
                    $this->writeShapeTxt($objWriter, $shape);
                }
            }
            
            $objWriter->endElement();
        }
    }

    /**
     * Write style of a slide
     * @param XMLWriter $objWriter
     * @param Slide $slide
     * @param int $incPage
     */
    public function writeStyleSlide(XMLWriter $objWriter, Slide $slide, $incPage)
    {
        // style:style
        $objWriter->startElement('style:style');
        $objWriter->writeAttribute('style:family', 'drawing-page');
        $objWriter->writeAttribute('style:name', 'stylePage'.$incPage);
        // style:style/style:drawing-page-properties
        $objWriter->startElement('style:drawing-page-properties');
        if (!is_null($oTransition = $slide->getTransition())) {
            $objWriter->writeAttribute('presentation:duration', 'PT'.number_format($oTransition->getAdvanceTimeTrigger() / 1000, 6, '.', '').'S');
            if ($oTransition->hasManualTrigger()) {
                $objWriter->writeAttribute('presentation:transition-type', 'manual');
            } elseif ($oTransition->hasTimeTrigger()) {
                $objWriter->writeAttribute('presentation:transition-type', 'automatic');
            }
            switch ($oTransition->getSpeed()) {
                case Transition::SPEED_FAST:
                    $objWriter->writeAttribute('presentation:transition-speed', 'fast');
                    break;
                case Transition::SPEED_MEDIUM:
                    $objWriter->writeAttribute('presentation:transition-speed', 'medium');
                    break;
                case Transition::SPEED_SLOW:
                    $objWriter->writeAttribute('presentation:transition-speed', 'slow');
                    break;
            }

            /**
             * http://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part1.html#property-presentation_transition-style
             */
            switch ($oTransition->getTransitionType()) {
                case Transition::TRANSITION_BLINDS_HORIZONTAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'horizontal-stripes');
                    break;
                case Transition::TRANSITION_BLINDS_VERTICAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'vertical-stripes');
                    break;
                case Transition::TRANSITION_CHECKER_HORIZONTAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'horizontal-checkerboard');
                    break;
                case Transition::TRANSITION_CHECKER_VERTICAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'vertical-checkerboard');
                    break;
                case Transition::TRANSITION_CIRCLE_HORIZONTAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_CIRCLE_VERTICAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_COMB_HORIZONTAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_COMB_VERTICAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_COVER_DOWN:
                    $objWriter->writeAttribute('presentation:transition-style', 'uncover-to-bottom');
                    break;
                case Transition::TRANSITION_COVER_LEFT:
                    $objWriter->writeAttribute('presentation:transition-style', 'uncover-to-left');
                    break;
                case Transition::TRANSITION_COVER_LEFT_DOWN:
                    $objWriter->writeAttribute('presentation:transition-style', 'uncover-to-lowerleft');
                    break;
                case Transition::TRANSITION_COVER_LEFT_UP:
                    $objWriter->writeAttribute('presentation:transition-style', 'uncover-to-upperleft');
                    break;
                case Transition::TRANSITION_COVER_RIGHT:
                    $objWriter->writeAttribute('presentation:transition-style', 'uncover-to-right');
                    break;
                case Transition::TRANSITION_COVER_RIGHT_DOWN:
                    $objWriter->writeAttribute('presentation:transition-style', 'uncover-to-lowerright');
                    break;
                case Transition::TRANSITION_COVER_RIGHT_UP:
                    $objWriter->writeAttribute('presentation:transition-style', 'uncover-to-upperright');
                    break;
                case Transition::TRANSITION_COVER_UP:
                    $objWriter->writeAttribute('presentation:transition-style', 'uncover-to-top');
                    break;
                case Transition::TRANSITION_CUT:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_DIAMOND:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_DISSOLVE:
                    $objWriter->writeAttribute('presentation:transition-style', 'dissolve');
                    break;
                case Transition::TRANSITION_FADE:
                    $objWriter->writeAttribute('presentation:transition-style', 'fade-from-center');
                    break;
                case Transition::TRANSITION_NEWSFLASH:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_PLUS:
                    $objWriter->writeAttribute('presentation:transition-style', 'close');
                    break;
                case Transition::TRANSITION_PULL_DOWN:
                    $objWriter->writeAttribute('presentation:transition-style', 'stretch-from-bottom');
                    break;
                case Transition::TRANSITION_PULL_LEFT:
                    $objWriter->writeAttribute('presentation:transition-style', 'stretch-from-left');
                    break;
                case Transition::TRANSITION_PULL_RIGHT:
                    $objWriter->writeAttribute('presentation:transition-style', 'stretch-from-right');
                    break;
                case Transition::TRANSITION_PULL_UP:
                    $objWriter->writeAttribute('presentation:transition-style', 'stretch-from-top');
                    break;
                case Transition::TRANSITION_PUSH_DOWN:
                    $objWriter->writeAttribute('presentation:transition-style', 'roll-from-bottom');
                    break;
                case Transition::TRANSITION_PUSH_LEFT:
                    $objWriter->writeAttribute('presentation:transition-style', 'roll-from-left');
                    break;
                case Transition::TRANSITION_PUSH_RIGHT:
                    $objWriter->writeAttribute('presentation:transition-style', 'roll-from-right');
                    break;
                case Transition::TRANSITION_PUSH_UP:
                    $objWriter->writeAttribute('presentation:transition-style', 'roll-from-top');
                    break;
                case Transition::TRANSITION_RANDOM:
                    $objWriter->writeAttribute('presentation:transition-style', 'random');
                    break;
                case Transition::TRANSITION_RANDOMBAR_HORIZONTAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'horizontal-lines');
                    break;
                case Transition::TRANSITION_RANDOMBAR_VERTICAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'vertical-lines');
                    break;
                case Transition::TRANSITION_SPLIT_IN_HORIZONTAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'close-horizontal');
                    break;
                case Transition::TRANSITION_SPLIT_OUT_HORIZONTAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'open-horizontal');
                    break;
                case Transition::TRANSITION_SPLIT_IN_VERTICAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'close-vertical');
                    break;
                case Transition::TRANSITION_SPLIT_OUT_VERTICAL:
                    $objWriter->writeAttribute('presentation:transition-style', 'open-vertical');
                    break;
                case Transition::TRANSITION_STRIPS_LEFT_DOWN:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_STRIPS_LEFT_UP:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_STRIPS_RIGHT_DOWN:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_STRIPS_RIGHT_UP:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_WEDGE:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_WIPE_DOWN:
                    $objWriter->writeAttribute('presentation:transition-style', 'fade-from-bottom');
                    break;
                case Transition::TRANSITION_WIPE_LEFT:
                    $objWriter->writeAttribute('presentation:transition-style', 'fade-from-left');
                    break;
                case Transition::TRANSITION_WIPE_RIGHT:
                    $objWriter->writeAttribute('presentation:transition-style', 'fade-from-right');
                    break;
                case Transition::TRANSITION_WIPE_UP:
                    $objWriter->writeAttribute('presentation:transition-style', 'fade-from-top');
                    break;
                case Transition::TRANSITION_ZOOM_IN:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
                case Transition::TRANSITION_ZOOM_OUT:
                    $objWriter->writeAttribute('presentation:transition-style', 'none');
                    break;
            }
        }
        $objWriter->endElement();
        // > style:style
        $objWriter->endElement();
    }
    

    /**
     * @param XMLWriter $objWriter
     * @param Shadow $oShadow
     * @todo Improve for supporting any direction (https://sinepost.wordpress.com/2012/02/16/theyve-got-atan-you-want-atan2/)
     */
    protected function writeStylePartShadow(XMLWriter $objWriter, Shadow $oShadow)
    {
        $objWriter->writeAttribute('draw:shadow', 'visible');
        $objWriter->writeAttribute('draw:shadow-color', '#' . $oShadow->getColor()->getRGB());
        if ($oShadow->getDirection() == 0 || $oShadow->getDirection() == 360) {
            $objWriter->writeAttribute('draw:shadow-offset-x', CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '0cm');
        } elseif ($oShadow->getDirection() == 45) {
            $objWriter->writeAttribute('draw:shadow-offset-x', CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
        } elseif ($oShadow->getDirection() == 90) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '0cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
        } elseif ($oShadow->getDirection() == 135) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '-' . CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
        } elseif ($oShadow->getDirection() == 180) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '-' . CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '0cm');
        } elseif ($oShadow->getDirection() == 225) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '-' . CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '-' . CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
        } elseif ($oShadow->getDirection() == 270) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '0cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '-' . CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
        } elseif ($oShadow->getDirection() == 315) {
            $objWriter->writeAttribute('draw:shadow-offset-x', CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '-' . CommonDrawing::pixelsToCentimeters($oShadow->getDistance()) . 'cm');
        }
        $objWriter->writeAttribute('draw:shadow-opacity', (100 - $oShadow->getAlpha()) . '%');
        $objWriter->writeAttribute('style:mirror', 'none');
    
    }
}
