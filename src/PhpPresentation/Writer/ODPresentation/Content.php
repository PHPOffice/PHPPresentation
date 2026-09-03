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

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PhpOffice\PhpPresentation\Shape\RichText\Field;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Note;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Shadow;

class Content extends AbstractDecoratorWriter
{
    /**
     * The prefix of a generated automatic style name, by family.
     *
     * These are the prefixes LibreOffice generates, so that a file this writer produces is named
     * the way one LibreOffice produces is.
     */
    private const STYLE_PREFIX = [
        'paragraph' => 'P',
        'text' => 'T',
        'graphic' => 'gr',
        'drawing-page' => 'dp',
        'table-row' => 'ro',
        'table-cell' => 'ce',
    ];

    /**
     * The element each kind of field writes its text with, where OpenDocument has one.
     *
     * The dated formats are not here: they are resolved by their prefix, since OOXML numbers
     * fourteen of them and OpenDocument has two elements for the lot.
     */
    private const FIELD_ODF = [
        Field::TYPE_SLIDENUM => 'text:page-number',
        Field::TYPE_SLIDECOUNT => 'text:page-count',
        'author' => 'text:author-name',
        'file' => 'text:file-name',
        'file1' => 'text:file-name',
        'file2' => 'text:file-name',
        'file3' => 'text:file-name',
    ];

    /**
     * Stores bullet styles for text shapes that include lists.
     *
     * @var array<string, array<string, mixed>>
     */
    protected $arrStyleBullet = [];

    /**
     * The automatic styles to write, by the name each was given.
     *
     * @var array<string, array{family: string, write: callable(XMLWriter): void}>
     */
    protected $automaticStyles = [];

    /**
     * The name each automatic style was given, by what that style writes.
     *
     * @var array<string, string>
     */
    protected $automaticStyleNames = [];

    /**
     * The name of the automatic style each styled object wears, by object.
     *
     * A generated name follows from the order the styles were added, so the pass that writes the
     * body cannot recompute it the way it used to recompute a hash. It reads it from here instead.
     *
     * @var array<int, string>
     */
    protected $automaticStyleNameByObject = [];

    /**
     * How many names of each family have been generated.
     *
     * @var array<string, int>
     */
    protected $automaticStyleCounters = [];

    /**
     * Used to track the current shape ID.
     *
     * @var int
     */
    protected $shapeId;

    public function render(): ZipInterface
    {
        $this->getZip()->addFromString('content.xml', $this->writeContent());

        return $this->getZip();
    }

    /**
     * Write content file to XML format.
     *
     * @return string XML Output
     */
    protected function writeContent(): string
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
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
        $objWriter->writeAttribute('xmlns:officeooo', 'http://openoffice.org/2009/office');
        $objWriter->writeAttribute('xmlns:loext', 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0');
        $objWriter->writeAttribute('office:version', '1.2');

        // office:automatic-styles
        $objWriter->startElement('office:automatic-styles');

        $this->shapeId = 0;
        $incSlide = 0;
        foreach ($this->getPresentation()->getAllSlides() as $pSlide) {
            // Slides
            $this->addSlideStyle($pSlide, $incSlide);

            // Shapes
            $this->addShapeStyles($objWriter, $pSlide->getShapeCollection());

            // The shapes of a slide note are written by writeSlideNote(), which increments the same
            // shape counter and names the styles of their text -- so they have to be collected here
            // too, or the note references three styles this pass never defines and the counter the
            // two passes share falls out of step.
            if ($pSlide->getNote() instanceof Note) {
                $this->addShapeStyles($objWriter, $pSlide->getNote()->getShapeCollection());
            }

            ++$incSlide;
        }
        // Style : Bullet
        if (!empty($this->arrStyleBullet)) {
            foreach ($this->arrStyleBullet as $key => $item) {
                $oStyle = $item['oStyle'];
                $arrLevel = explode(';', $item['level']);
                // style:style
                $objWriter->startElement('text:list-style');
                $objWriter->writeAttribute('style:name', 'L_' . $key);
                foreach ($arrLevel as $level) {
                    if ('' != $level) {
                        $oAlign = $item['oAlign_' . $level];
                        if (Bullet::TYPE_NUMERIC == $oStyle->getBulletType()) {
                            [$numFormat, $numPrefix, $numSuffix] = $this->getNumericBulletFormat($oStyle->getBulletNumericStyle());
                            // text:list-level-style-number
                            $objWriter->startElement('text:list-level-style-number');
                            $objWriter->writeAttribute('text:level', (int) $level + 1);
                            $objWriter->writeAttribute('style:num-format', $numFormat);
                            $objWriter->writeAttributeIf('' !== $numPrefix, 'style:num-prefix', $numPrefix);
                            $objWriter->writeAttributeIf('' !== $numSuffix, 'style:num-suffix', $numSuffix);
                            $objWriter->writeAttributeIf(1 != $oStyle->getBulletNumericStartAt(), 'text:start-value', $oStyle->getBulletNumericStartAt());
                        } else {
                            // text:list-level-style-bullet
                            $objWriter->startElement('text:list-level-style-bullet');
                            $objWriter->writeAttribute('text:level', (int) $level + 1);
                            $objWriter->writeAttribute('text:bullet-char', $oStyle->getBulletChar());
                        }
                        // style:list-level-properties
                        $objWriter->startElement('style:list-level-properties');
                        if ($oAlign->getIndent() < 0) {
                            $objWriter->writeAttribute('text:space-before', CommonDrawing::pixelsToCentimeters((int) ($oAlign->getMarginLeft() - (-1 * $oAlign->getIndent()))) . 'cm');
                            $objWriter->writeAttribute('text:min-label-width', CommonDrawing::pixelsToCentimeters((int) (-1 * $oAlign->getIndent())) . 'cm');
                        } else {
                            $objWriter->writeAttribute('text:space-before', (CommonDrawing::pixelsToCentimeters((int) ($oAlign->getMarginLeft() - $oAlign->getIndent()))) . 'cm');
                            $objWriter->writeAttribute('text:min-label-width', CommonDrawing::pixelsToCentimeters((int) $oAlign->getIndent()) . 'cm');
                        }

                        $objWriter->endElement();
                        // style:text-properties
                        $objWriter->startElement('style:text-properties');
                        $objWriter->writeAttribute('fo:font-family', $oStyle->getBulletFont());
                        $objWriter->writeAttribute('style:font-family-generic', 'swiss');
                        $objWriter->writeAttribute('style:use-window-font-color', 'true');
                        $objWriter->writeAttribute('fo:font-size', '100%');
                        $objWriter->endElement();
                        $objWriter->endElement();
                    }
                }
                $objWriter->endElement();
            }
        }
        // Emitted from the pool, in the order the styles were first needed -- which is the order
        // `office:automatic-styles` wants, since it precedes the content that references it.
        foreach ($this->automaticStyles as $styleName => $automaticStyle) {
            // style:style
            $objWriter->startElement('style:style');
            $objWriter->writeAttribute('style:name', $styleName);
            $objWriter->writeAttribute('style:family', $automaticStyle['family']);
            ($automaticStyle['write'])($objWriter);
            $objWriter->endElement();
        }
        $objWriter->endElement();

        //===============================================
        // Body
        //===============================================
        // office:body
        $objWriter->startElement('office:body');
        // office:body > office:presentation
        $objWriter->startElement('office:presentation');

        // Write slides
        $slideCount = $this->getPresentation()->getSlideCount();
        $this->shapeId = 0;
        for ($i = 0; $i < $slideCount; ++$i) {
            $pSlide = $this->getPresentation()->getSlide($i);
            $objWriter->startElement('draw:page');
            $objWriter->writeAttribute('draw:name', $this->getSlideName($pSlide, $i + 1));
            $objWriter->writeAttribute('draw:master-page-name', 'Standard');
            $objWriter->writeAttribute('draw:style-name', $this->getAutomaticStyleName($pSlide));
            // Shapes
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
                } elseif ($shape instanceof Media) {
                    $this->writeShapeMedia($objWriter, $shape);
                } elseif ($shape instanceof AbstractDrawingAdapter) {
                    $this->writeShapeDrawing($objWriter, $shape);
                } elseif ($shape instanceof Group) {
                    $this->writeShapeGroup($objWriter, $shape);
                } elseif ($shape instanceof Comment) {
                    $this->writeShapeComment($objWriter, $shape);
                }
            }
            // Slide Note
            if ($pSlide->getNote() instanceof Note) {
                $this->writeSlideNote($objWriter, $pSlide->getNote());
            }

            $objWriter->endElement();
        }

        // office:document-content > office:body > office:presentation > presentation:settings
        $objWriter->startElement('presentation:settings');
        if ($this->getPresentation()->getPresentationProperties()->isLoopContinuouslyUntilEsc()) {
            $objWriter->writeAttribute('presentation:endless', 'true');
            $objWriter->writeAttribute('presentation:pause', 'PT0S');
            $objWriter->writeAttribute('presentation:mouse-visible', 'false');
        }
        if ($this->getPresentation()->getPresentationProperties()->getSlideshowType() === PresentationProperties::SLIDESHOW_TYPE_BROWSE) {
            $objWriter->writeAttribute('presentation:full-screen', 'false');
        }
        $objWriter->endElement();

        // > office:presentation
        $objWriter->endElement();
        // > office:body
        $objWriter->endElement();
        // > office:document-content
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * The ODF number format each OOXML autonumber alphabet is written as.
     *
     * Measured 26 Aug 2026 by writing a deck holding all 41 schemes of `Style\Bullet` and letting
     * LibreOffice convert it: these are the formats it produced. An alphabet named nowhere here is
     * written as arabic, and those are exactly the eleven schemes LibreOffice drops the numbering
     * of altogether -- `arabic1Minus`, `arabic2Minus`, `arabicDbPeriod`, `arabicDbPlain`, the three
     * `hindi*` and `hindiAlpha1Period`, `ea1JpnChsDbPeriod`, `ea1JpnKorPeriod`, `ea1JpnKorPlain`.
     * A marker of the wrong alphabet is a smaller loss than a list that stops being one.
     *
     * @var array<string, string>
     */
    protected const NUMERIC_BULLET_FORMAT = [
        'alphaLc' => 'a',
        'alphaUc' => 'A',
        'romanLc' => 'i',
        'romanUc' => 'I',
        'circleNumDb' => "\u{2460}, \u{2461}, \u{2462}, ...",
        'circleNumWdBlack' => "\u{2460}, \u{2461}, \u{2462}, ...",
        'circleNumWdWhite' => "\u{2460}, \u{2461}, \u{2462}, ...",
        'ea1Chs' => "\u{58F9}, \u{8D30}, \u{53C1}, ...",
        'ea1Cht' => "\u{58F9}, \u{8CB3}, \u{53C3}, ...",
        'hebrew2' => "\u{05D0}, \u{05D1}, \u{05D2}, ...",
        // LibreOffice writes the Thai letter sequence for both Thai schemes, the numeric one
        // included. Measured, not chosen.
        'thaiAlpha' => "\u{0E01}, \u{0E02}, \u{0E04}, ...",
        'thaiNum' => "\u{0E01}, \u{0E02}, \u{0E04}, ...",
    ];

    /**
     * The prefix and the suffix each OOXML autonumber separator is written as. ODF keeps the
     * punctuation of a marker apart from its number, so `arabicParenBoth` is a format and two
     * attributes rather than one token.
     *
     * @var array<string, array<int, string>>
     */
    protected const NUMERIC_BULLET_SEPARATOR = [
        'Period' => ['', '.'],
        'ParenBoth' => ['(', ')'],
        'ParenR' => ['', ')'],
        'Minus' => ['', '-'],
        'Plain' => ['', ''],
    ];

    /**
     * The three parts an OOXML autonumber scheme is taken apart into: the format of the number,
     * then the prefix and the suffix that carry its punctuation.
     *
     * @return array<int, string>
     */
    protected function getNumericBulletFormat(string $scheme): array
    {
        if (!preg_match('/^(.*?)(Period|ParenBoth|ParenR|Minus|Plain)$/', $scheme, $matches)) {
            return ['1', '', ''];
        }

        [$prefix, $suffix] = self::NUMERIC_BULLET_SEPARATOR[$matches[2]];

        return [self::NUMERIC_BULLET_FORMAT[$matches[1]] ?? '1', $prefix, $suffix];
    }

    /**
     * The name of the `text:list-style` the marker of a paragraph is written under, or an empty
     * string when the paragraph asks for no marker and is written outside any list.
     *
     * A bullet list and a numbered list are the same `text:list` in ODF and differ only in the
     * style it names, so this one name answers both what to write and where a list ends: two
     * paragraphs belong to the same list exactly when they name the same style.
     */
    protected function getListStyleName(Paragraph $paragraph): string
    {
        $bulletType = $paragraph->getBulletStyle()->getBulletType();
        if (Bullet::TYPE_BULLET != $bulletType && Bullet::TYPE_NUMERIC != $bulletType) {
            return '';
        }

        return 'L_' . $paragraph->getBulletStyle()->getHashCode();
    }

    /**
     * Write the decorative flag of a shape, ie the shape is ignored by assistive
     * technologies. It has to be written before any child of the shape element.
     *
     * ODF has no such attribute before version 1.4, so the LibreOffice extension is used.
     */
    protected function writeShapeDecorative(XMLWriter $objWriter, AbstractShape $shape): void
    {
        if (!$shape->isDecorative()) {
            return;
        }

        $objWriter->writeAttribute('loext:decorative', 'true');
    }

    /**
     * Write the description of a shape, exposed to assistive technologies as the
     * alternative text. It has to be the first child of the shape element.
     */
    protected function writeShapeDescription(XMLWriter $objWriter, AbstractShape $shape): void
    {
        if ('' === $shape->getDescription()) {
            return;
        }

        $objWriter->writeElement('svg:desc', $shape->getDescription());
    }

    /**
     * Write the hyperlink a shape as a whole carries, which ODF states as a listener for the
     * click on it rather than as a property of the shape.
     *
     * Where the element it belongs to allows it is not the same for every shape: a `draw:frame`
     * takes it after its content and before `svg:desc`, while `draw:line` and `draw:g` take it
     * after `svg:desc` instead. The caller places it; this only writes it.
     */
    protected function writeShapeHyperlink(XMLWriter $objWriter, AbstractShape $shape): void
    {
        if (!$shape->hasHyperlink()) {
            return;
        }

        // office:event-listeners
        $objWriter->startElement('office:event-listeners');
        // presentation:event-listener
        $objWriter->startElement('presentation:event-listener');
        $objWriter->writeAttribute('script:event-name', 'dom:click');
        $objWriter->writeAttribute('presentation:action', 'show');
        $objWriter->writeAttribute('xlink:href', $this->getHyperlinkHref($shape->getHyperlink()));
        $objWriter->writeAttribute('xlink:type', 'simple');
        $objWriter->writeAttribute('xlink:show', 'embed');
        $objWriter->writeAttribute('xlink:actuate', 'onRequest');
        // > presentation:event-listener
        $objWriter->endElement();
        // > office:event-listeners
        $objWriter->endElement();
    }

    /**
     * Write picture.
     */
    protected function writeShapeMedia(XMLWriter $objWriter, Media $shape): void
    {
        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('draw:name', $shape->getName());
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getWidth()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getHeight()), 3) . 'cm');
        $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('draw:style-name', $this->getAutomaticStyleName($shape));
        $this->writeShapeDecorative($objWriter, $shape);
        // draw:frame > draw:plugin
        $objWriter->startElement('draw:plugin');
        $objWriter->writeAttribute('xlink:href', 'Pictures/' . $this->writtenPart($shape)->getIndexedFilename());
        $objWriter->writeAttribute('xlink:type', 'simple');
        $objWriter->writeAttribute('xlink:show', 'embed');
        $objWriter->writeAttribute('xlink:actuate', 'onLoad');
        $objWriter->writeAttribute('draw:mime-type', 'application/vnd.sun.star.media');

        $objWriter->startElement('draw:param');
        $objWriter->writeAttribute('draw:name', 'Loop');
        $objWriter->writeAttribute('draw:value', 'false');
        $objWriter->endElement();
        $objWriter->startElement('draw:param');
        $objWriter->writeAttribute('draw:name', 'Mute');
        $objWriter->writeAttribute('draw:value', 'false');
        $objWriter->endElement();
        $objWriter->startElement('draw:param');
        $objWriter->writeAttribute('draw:name', 'VolumeDB');
        $objWriter->writeAttribute('draw:value', 0);
        $objWriter->endElement();
        $objWriter->startElement('draw:param');
        $objWriter->writeAttribute('draw:name', 'Zoom');
        $objWriter->writeAttribute('draw:value', 'fit');
        $objWriter->endElement();

        // draw:frame > ## draw:plugin
        $objWriter->endElement();

        $this->writeShapeHyperlink($objWriter, $shape);
        $this->writeShapeDescription($objWriter, $shape);

        // ## draw:frame
        $objWriter->endElement();
    }

    /**
     * Write picture.
     */
    protected function writeShapeDrawing(XMLWriter $objWriter, AbstractDrawingAdapter $shape): void
    {
        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('draw:name', $shape->getName());
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getWidth()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getHeight()), 3) . 'cm');
        $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('draw:style-name', $this->getAutomaticStyleName($shape));
        $this->writeShapeDecorative($objWriter, $shape);
        // draw:image
        $objWriter->startElement('draw:image');
        $objWriter->writeAttribute('xlink:href', 'Pictures/' . $this->writtenPart($shape)->getIndexedFilename());
        $objWriter->writeAttribute('xlink:type', 'simple');
        $objWriter->writeAttribute('xlink:show', 'embed');
        $objWriter->writeAttribute('xlink:actuate', 'onLoad');
        $objWriter->writeAttribute('loext:mime-type', $shape->getMimeType());
        $objWriter->writeElement('text:p');
        $objWriter->endElement();

        $this->writeShapeHyperlink($objWriter, $shape);
        $this->writeShapeDescription($objWriter, $shape);

        $objWriter->endElement();
    }

    /**
     * Write text.
     */
    protected function writeShapeTxt(XMLWriter $objWriter, RichText $shape): void
    {
        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('draw:style-name', $this->getAutomaticStyleName($shape));
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getWidth()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getHeight()), 3) . 'cm');
        if ($shape->getRotation() != 0) {
            $rotRad = deg2rad($shape->getRotation());

            $translateX = Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetY()), 3) . 'cm';
            $translateY = Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetX()), 3) . 'cm';
            $objWriter->writeAttribute(
                'draw:transform',
                'rotate (-' . $rotRad . ') translate (' . $translateX . ' ' . $translateY . ')'
            );
        } else {
            $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetX()), 3) . 'cm');
            $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetY()), 3) . 'cm');
        }
        $this->writeShapeDecorative($objWriter, $shape);
        // draw:text-box
        $objWriter->startElement('draw:text-box');

        $paragraphs = $shape->getParagraphs();
        $paragraphId = 0;
        $sCstShpOpenList = '';
        $iCstShpLastBulletLvl = 0;
        $fieldName = $this->getPlaceholderField($shape);

        foreach ($paragraphs as $paragraph) {
            $sCstShpListStyle = $this->getListStyleName($paragraph);
            // Close the open list, when this paragraph is not part of it. It is the paragraph
            // being written that says so, not the one before it: asking the one before left a
            // plain paragraph that follows a list inside the last item of that list.
            if ('' !== $sCstShpOpenList && $sCstShpListStyle !== $sCstShpOpenList) {
                for ($iInc = $iCstShpLastBulletLvl; $iInc >= 0; --$iInc) {
                    // text:list-item
                    $objWriter->endElement();
                    // text:list
                    $objWriter->endElement();
                }
                $sCstShpOpenList = '';
            }
            //===============================================
            // Paragraph
            //===============================================
            if ('' === $sCstShpListStyle) {
                ++$paragraphId;
                // text:p
                $objWriter->startElement('text:p');
                $objWriter->writeAttribute('text:style-name', $this->getAutomaticStyleName($paragraph));

                // Loop trough rich text elements
                $richtexts = $paragraph->getRichTextElements();
                $richtextId = 0;
                foreach ($richtexts as $richtext) {
                    ++$richtextId;
                    if ($richtext instanceof TextElement) {
                        // text:span
                        $objWriter->startElement('text:span');
                        if ($richtext instanceof Run) {
                            $objWriter->writeAttribute('text:style-name', $this->getAutomaticStyleName($richtext));
                        }
                        if (true === $richtext->hasHyperlink() && '' != $richtext->getHyperlink()->getUrl()) {
                            // text:a
                            $objWriter->startElement('text:a');
                            $objWriter->writeAttribute('xlink:type', 'simple');
                            $objWriter->writeAttribute('xlink:href', $this->getHyperlinkHref($richtext->getHyperlink()));
                            $objWriter->text($richtext->getText());
                            $objWriter->endElement();
                        } elseif (null !== ($field = $this->getFieldElement($richtext, $fieldName))) {
                            $objWriter->writeElement($field, $richtext->getText());
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
                    }
                }
                $objWriter->endElement();
                //===============================================
                // Bullet list
                //===============================================
            } else {
                // Open the list
                if ('' === $sCstShpOpenList || $iCstShpLastBulletLvl < $paragraph->getAlignment()->getLevel()) {
                    // text:list
                    $objWriter->startElement('text:list');
                    $objWriter->writeAttribute('text:style-name', $sCstShpListStyle);
                }
                if ('' !== $sCstShpOpenList) {
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
                $objWriter->writeAttribute('text:style-name', $this->getAutomaticStyleName($paragraph));

                // Loop trough rich text elements
                $richtexts = $paragraph->getRichTextElements();
                $richtextId = 0;
                foreach ($richtexts as $richtext) {
                    ++$richtextId;
                    if ($richtext instanceof TextElement) {
                        // text:span
                        $objWriter->startElement('text:span');
                        if ($richtext instanceof Run) {
                            $objWriter->writeAttribute('text:style-name', $this->getAutomaticStyleName($richtext));
                        }
                        if (true === $richtext->hasHyperlink() && '' != $richtext->getHyperlink()->getUrl()) {
                            // text:a
                            $objWriter->startElement('text:a');
                            $objWriter->writeAttribute('xlink:type', 'simple');
                            $objWriter->writeAttribute('xlink:href', $this->getHyperlinkHref($richtext->getHyperlink()));
                            $objWriter->text($richtext->getText());
                            $objWriter->endElement();
                        } elseif (null !== ($field = $this->getFieldElement($richtext, $fieldName))) {
                            $objWriter->writeElement($field, $richtext->getText());
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
                    }
                }
                $objWriter->endElement();
            }
            $sCstShpOpenList = $sCstShpListStyle;
            $iCstShpLastBulletLvl = $paragraph->getAlignment()->getLevel();
        }

        // Close the open list
        if ('' !== $sCstShpOpenList) {
            for ($iInc = $iCstShpLastBulletLvl; $iInc >= 0; --$iInc) {
                // text:list-item
                $objWriter->endElement();
                // text:list
                $objWriter->endElement();
            }
        }

        // > draw:text-box
        $objWriter->endElement();

        $this->writeShapeHyperlink($objWriter, $shape);
        $this->writeShapeDescription($objWriter, $shape);

        // > draw:frame
        $objWriter->endElement();
    }

    /**
     * The element a field writes its text with.
     *
     * OpenDocument names a field by what it is rather than by how it is formatted, so the
     * fourteen dated formats OOXML numbers all say "date" here -- save the four that carry no
     * date at all -- and the format itself would travel as a data style. A field OpenDocument
     * has no element for is written as the text it stands in for, which is what an application
     * that cannot compute it would show anyway.
     */
    protected function getFieldElement(TextElement $richtext, ?string $placeholderField): ?string
    {
        if (!$richtext instanceof Field) {
            return $placeholderField;
        }

        $type = $richtext->getType();
        if (0 === strpos($type, 'datetime')) {
            return in_array($type, ['datetime10', 'datetime11', 'datetime12', 'datetime13'], true)
                ? 'text:time'
                : 'text:date';
        }

        return self::FIELD_ODF[$type] ?? null;
    }

    /**
     * The element a placeholder writes its text with, if that text is a field.
     *
     * The number of a slide and its date are not the text the shape was given: they are
     * fields the reader fills in, and the text is only what they stand in for.
     */
    protected function getPlaceholderField(RichText $shape): ?string
    {
        $placeholder = $shape->getPlaceholder();
        if (null === $placeholder) {
            return null;
        }

        switch ($placeholder->getType()) {
            case Placeholder::PH_TYPE_SLIDENUM:
                return 'text:page-number';
            case Placeholder::PH_TYPE_DATETIME:
                return 'text:date';
            default:
                return null;
        }
    }

    /**
     * Write Comment.
     */
    protected function writeShapeComment(XMLWriter $objWriter, Comment $oShape): void
    {
        // Note : This element is not valid in the Schema 1.2
        // officeooo:annotation
        $objWriter->startElement('officeooo:annotation');
        $objWriter->writeAttribute('svg:x', number_format(CommonDrawing::pixelsToCentimeters((int) $oShape->getOffsetX()), 2, '.', '') . 'cm');
        $objWriter->writeAttribute('svg:y', number_format(CommonDrawing::pixelsToCentimeters((int) $oShape->getOffsetY()), 2, '.', '') . 'cm');

        if ($oShape->getAuthor() instanceof Comment\Author) {
            $objWriter->writeElement('dc:creator', $oShape->getAuthor()->getName());
        }
        $objWriter->writeElement('dc:date', date('Y-m-d\TH:i:s', $oShape->getDate()));
        $objWriter->writeElement('text:p', $oShape->getText());

        // ## officeooo:annotation
        $objWriter->endElement();
    }

    protected function writeShapeLine(XMLWriter $objWriter, Line $shape): void
    {
        // draw:line
        $objWriter->startElement('draw:line');
        $objWriter->writeAttribute('draw:style-name', $this->getAutomaticStyleName($shape));
        $objWriter->writeAttribute('svg:x1', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y1', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('svg:x2', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetX() + $shape->getWidth()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y2', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetY() + $shape->getHeight()), 3) . 'cm');

        $this->writeShapeDecorative($objWriter, $shape);
        $this->writeShapeDescription($objWriter, $shape);
        $this->writeShapeHyperlink($objWriter, $shape);

        // text:p
        $objWriter->writeElement('text:p');

        $objWriter->endElement();
    }

    /**
     * The name ODF addresses a slide by, which every page carries so that a link can reach it.
     *
     * A slide with no name of its own is named after its position. Not as `page3`, though:
     * LibreOffice generates exactly that form for the pages it exports and treats it as its own,
     * so a link to `#page3` is left as an unresolved URI rather than followed.
     */
    protected function getSlideName(Slide $slide, int $slideNumber): string
    {
        return $slide->getName() ?? 'Slide ' . $slideNumber;
    }

    /**
     * The target of a hyperlink, in the terms ODF addresses it.
     *
     * A link to another slide is stored as the PowerPoint action string `ppaction://hlinksldjump`,
     * which ODF has never heard of -- there a slide is addressed by name.
     */
    protected function getHyperlinkHref(Hyperlink $hyperlink): string
    {
        if (!$hyperlink->isInternal()) {
            return $hyperlink->getUrl();
        }

        $slideNumber = $hyperlink->getSlideNumber();
        if ($slideNumber < 1 || $slideNumber > $this->getPresentation()->getSlideCount()) {
            return $hyperlink->getUrl();
        }

        return '#' . $this->getSlideName($this->getPresentation()->getSlide($slideNumber - 1), $slideNumber);
    }

    /**
     * Write table Shape.
     */
    protected function writeShapeTable(XMLWriter $objWriter, Table $shape): void
    {
        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getHeight()), 3) . 'cm');
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getWidth()), 3) . 'cm');

        $this->writeShapeDecorative($objWriter, $shape);

        $arrayRows = $shape->getRows();
        if (!empty($arrayRows)) {
            $firstRow = reset($arrayRows);
            $arrayCells = $firstRow->getCells();
            // table:table
            $objWriter->startElement('table:table');
            // A table on a drawing page says which of its rows are styled apart with the two flags
            // below, which are what `firstRow` and `bandRow` say on `a:tblPr`. `table:table-header-rows`
            // is the other ODF table model, the one a text table uses to repeat a header across
            // pages, and a consumer reading a drawing table drops the rows it wraps.
            $objWriter->writeAttributeIf($shape->isFirstRow(), 'table:use-first-row-styles', 'true');
            $objWriter->writeAttributeIf($shape->isBandRow(), 'table:use-banding-rows-styles', 'true');
            foreach ($arrayCells as $shapeCell) {
                $objWriter->startElement('table:table-column');
                $objWriter->endElement();
            }
            foreach ($arrayRows as $shapeRow) {
                // table:table-row
                $objWriter->startElement('table:table-row');
                $objWriter->writeAttribute('table:style-name', $this->getAutomaticStyleName($shapeRow));
                //@todo getFill

                $numColspan = 0;
                foreach ($shapeRow->getCells() as $shapeCell) {
                    if (0 == $numColspan) {
                        // table:table-cell
                        $objWriter->startElement('table:table-cell');
                        $objWriter->writeAttribute('table:style-name', $this->getAutomaticStyleName($shapeCell));
                        if ($shapeCell->getColspan() > 1) {
                            $objWriter->writeAttribute('table:number-columns-spanned', $shapeCell->getColspan());
                            $numColspan = $shapeCell->getColspan() - 1;
                        }

                        // text:p
                        $objWriter->startElement('text:p');

                        // text:span
                        foreach ($shapeCell->getParagraphs() as $shapeParagraph) {
                            foreach ($shapeParagraph->getRichTextElements() as $shapeRichText) {
                                if ($shapeRichText instanceof TextElement) {
                                    // text:span
                                    $objWriter->startElement('text:span');
                                    if ($shapeRichText instanceof Run) {
                                        $objWriter->writeAttribute('text:style-name', $this->getAutomaticStyleName($shapeRichText));
                                    }
                                    if (true === $shapeRichText->hasHyperlink() && '' !== $shapeRichText->getHyperlink()->getUrl()) {
                                        // text:a
                                        $objWriter->startElement('text:a');
                                        $objWriter->writeAttribute('xlink:type', 'simple');
                                        $objWriter->writeAttribute('xlink:href', $this->getHyperlinkHref($shapeRichText->getHyperlink()));
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
                        --$numColspan;
                    }
                }
                // > table:table-row
                $objWriter->endElement();
            }
            // > table:table
            $objWriter->endElement();
        }

        $this->writeShapeHyperlink($objWriter, $shape);
        $this->writeShapeDescription($objWriter, $shape);

        // > draw:frame
        $objWriter->endElement();
    }

    /**
     * Write table Chart.
     */
    protected function writeShapeChart(XMLWriter $objWriter, Chart $shape): void
    {
        $arrayChart = $this->getArrayChart();
        $arrayChart[$this->shapeId] = $shape;
        $this->setArrayChart($arrayChart);

        // draw:frame
        $objWriter->startElement('draw:frame');
        $objWriter->writeAttribute('draw:name', $shape->getTitle()->getText());
        $objWriter->writeAttribute('svg:x', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetX()), 3) . 'cm');
        $objWriter->writeAttribute('svg:y', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getOffsetY()), 3) . 'cm');
        $objWriter->writeAttribute('svg:height', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getHeight()), 3) . 'cm');
        $objWriter->writeAttribute('svg:width', Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $shape->getWidth()), 3) . 'cm');

        $this->writeShapeDecorative($objWriter, $shape);

        // draw:object
        $objWriter->startElement('draw:object');
        $objWriter->writeAttribute('xlink:href', './Object ' . $this->shapeId);
        $objWriter->writeAttribute('xlink:type', 'simple');
        $objWriter->writeAttribute('xlink:show', 'embed');

        // > draw:object
        $objWriter->endElement();

        $this->writeShapeHyperlink($objWriter, $shape);
        $this->writeShapeDescription($objWriter, $shape);

        // > draw:frame
        $objWriter->endElement();
    }

    /**
     * Writes a group of shapes.
     */
    protected function writeShapeGroup(XMLWriter $objWriter, Group $group): void
    {
        // draw:g
        $objWriter->startElement('draw:g');

        $this->writeShapeDecorative($objWriter, $group);
        $this->writeShapeDescription($objWriter, $group);
        $this->writeShapeHyperlink($objWriter, $group);

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
            } elseif ($shape instanceof AbstractDrawingAdapter) {
                $this->writeShapeDrawing($objWriter, $shape);
            } elseif ($shape instanceof Group) {
                $this->writeShapeGroup($objWriter, $shape);
            }
        }

        $objWriter->endElement(); // draw:g
    }

    /**
     * Name the automatic style every shape of a collection wears, walking a group the way the pass
     * that writes the body walks one -- nested groups included, which writeShapeGroup() does and
     * collecting the styles used not to.
     *
     * @param iterable<AbstractShape> $shapes
     */
    protected function addShapeStyles(XMLWriter $objWriter, iterable $shapes): void
    {
        foreach ($shapes as $shape) {
            // Increment $this->shapeId
            ++$this->shapeId;

            // Check type
            if ($shape instanceof RichText) {
                $this->addTxtStyle($shape);
            }
            if ($shape instanceof AbstractDrawingAdapter) {
                $this->addDrawingStyle($shape);
            }
            if ($shape instanceof Line) {
                $this->addLineStyle($shape);
            }
            if ($shape instanceof Table) {
                $this->addTableStyle($shape);
            }
            // A group inside a group is walked by writeShapeGroup(), so it has to be walked here too
            if ($shape instanceof Group) {
                $this->addShapeStyles($objWriter, $shape->getShapeCollection());
            }
        }
    }

    /**
     * Name the automatic style a paragraph wears, sharing one with every paragraph that writes the
     * same thing.
     */
    protected function addParagraphStyle(Paragraph $paragraph): string
    {
        return $this->shareAutomaticStyle('paragraph', function (XMLWriter $objWriter) use ($paragraph): void {
            $this->writeParagraphStyleBody($objWriter, $paragraph);
        }, $paragraph);
    }

    /**
     * Name the automatic style a run wears, sharing one with every run that writes the same thing.
     */
    protected function addTextStyle(Run $run): string
    {
        return $this->shareAutomaticStyle('text', function (XMLWriter $objWriter) use ($run): void {
            $this->writeTextStyleBody($objWriter, $run);
        }, $run);
    }

    /**
     * Give a style a name, reusing the name of an identical one.
     *
     * The key is the family and the body the style writes -- not `getHashCode()`, which is the
     * identity of the *content*: two paragraphs styled alike but holding different text hash
     * differently, which is every paragraph of a real document.
     *
     * @param string                   $family    the ODF style family
     * @param callable(XMLWriter): void $writeBody writes everything inside `style:style` -- which is
     *                                            what makes two styles the same, and what the pass
     *                                            that writes the definitions replays
     * @param object                   $owner     the object wearing the style, which the second pass
     *                                            looks the name up by
     */
    private function shareAutomaticStyle(string $family, callable $writeBody, object $owner): string
    {
        $bodyWriter = new XMLWriter();
        $writeBody($bodyWriter);

        $key = $family . "\0" . $bodyWriter->getData();
        if (!isset($this->automaticStyleNames[$key])) {
            $counter = ($this->automaticStyleCounters[$family] ?? 0) + 1;
            $this->automaticStyleCounters[$family] = $counter;
            $styleName = self::STYLE_PREFIX[$family] . $counter;
            $this->automaticStyleNames[$key] = $styleName;
            $this->automaticStyles[$styleName] = [
                'family' => $family,
                'write' => $writeBody,
            ];
        }

        return $this->automaticStyleNameByObject[spl_object_id($owner)] = $this->automaticStyleNames[$key];
    }

    /**
     * The name of the automatic style an object was given while the styles were collected.
     */
    private function getAutomaticStyleName(object $owner): string
    {
        return $this->automaticStyleNameByObject[spl_object_id($owner)];
    }

    /**
     * @param Paragraph|Run $item
     */
    protected function writeParagraphStyleBody(XMLWriter $objWriter, $item): void
    {
        // style:paragraph-properties
        $objWriter->startElement('style:paragraph-properties');
        $objWriter->writeAttributeIf(
            $item->getLineSpacingMode() === Paragraph::LINE_SPACING_MODE_PERCENT,
            'fo:line-height',
            $item->getLineSpacing() . '%'
        );
        $objWriter->writeAttributeIf(
            $item->getLineSpacingMode() === Paragraph::LINE_SPACING_MODE_POINT,
            'fo:line-height',
            $item->getLineSpacing() . 'pt'
        );
        // six decimals rather than three: a point is 127/3600 cm, and three decimals lose about a
        // seventieth of a point of the spacing
        $objWriter->writeAttribute(
            'fo:margin-top',
            round(CommonDrawing::pointstoCentimeters($item->getSpacingBefore()), 6) . 'cm'
        );
        $objWriter->writeAttribute(
            'fo:margin-bottom',
            round(CommonDrawing::pointstoCentimeters($item->getSpacingAfter()), 6) . 'cm'
        );
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
        $objWriter->writeAttribute(
            'style:writing-mode',
            $item->getAlignment()->isRTL() ? 'rl-tb' : 'lr-tb'
        );
        $objWriter->endElement();
    }

    /**
     * @param Paragraph|Run $item
     */
    protected function writeTextStyleBody(XMLWriter $objWriter, $item): void
    {
        // style:style > style:text-properties
        $objWriter->startElement('style:text-properties');
        $objWriter->writeAttribute('fo:color', '#' . $item->getFont()->getColor()->getRGB());
        switch ($item->getFont()->getCapitalization()) {
            case Font::CAPITALIZATION_NONE:
                $objWriter->writeAttribute('fo:text-transform', 'none');

                break;
            case Font::CAPITALIZATION_ALL:
                $objWriter->writeAttribute('fo:text-transform', 'uppercase');

                break;
            case Font::CAPITALIZATION_SMALL:
                $objWriter->writeAttribute('fo:text-transform', 'lowercase');

                break;
        }
        switch ($item->getFont()->getFormat()) {
            case Font::FORMAT_LATIN:
                $objWriter->writeAttribute('fo:font-family', $item->getFont()->getName());
                $objWriter->writeAttribute('fo:font-size', $item->getFont()->getSize() . 'pt');
                $objWriter->writeAttributeIf($item->getFont()->isBold(), 'fo:font-weight', 'bold');
                $objWriter->writeAttributeIf($item->getFont()->isItalic(), 'fo:font-style', 'italic');
                $objWriter->writeAttribute('fo:language', ($item->getLanguage() ? substr($item->getLanguage(), 0, 2) : 'en'));
                $objWriter->writeAttribute('style:script-type', 'latin');

                break;
            case Font::FORMAT_EAST_ASIAN:
                $objWriter->writeAttribute('style:font-family-asian', $item->getFont()->getName());
                $objWriter->writeAttribute('style:font-size-asian', $item->getFont()->getSize() . 'pt');
                $objWriter->writeAttributeIf($item->getFont()->isBold(), 'style:font-weight-asian', 'bold');
                $objWriter->writeAttributeIf($item->getFont()->isItalic(), 'style:font-style-asian', 'italic');
                $objWriter->writeAttribute('style:language-asian', ($item->getLanguage() ? $item->getLanguage() : 'en'));
                $objWriter->writeAttribute('style:script-type', 'asian');

                break;
            case Font::FORMAT_COMPLEX_SCRIPT:
                $objWriter->writeAttribute('style:font-family-complex', $item->getFont()->getName());
                $objWriter->writeAttribute('style:font-size-complex', $item->getFont()->getSize() . 'pt');
                $objWriter->writeAttributeIf($item->getFont()->isBold(), 'style:font-weight-complex', 'bold');
                $objWriter->writeAttributeIf($item->getFont()->isItalic(), 'style:font-style-complex', 'italic');
                $objWriter->writeAttribute('style:language-complex', ($item->getLanguage() ? $item->getLanguage() : 'en'));
                $objWriter->writeAttribute('style:script-type', 'complex');

                break;
        }
        // the underline and the strikethrough are one family for the whole run, whatever
        // the script the rest of it was spelled for
        $this->writeFontStates($objWriter, $item->getFont());

        // > style:style > style:text-properties
        $objWriter->endElement();
    }

    /**
     * Name the automatic style a text shape wears, and collect the styles of the text inside it.
     */
    protected function addTxtStyle(RichText $shape): void
    {
        $this->shareAutomaticStyle('graphic', function (XMLWriter $objWriter) use ($shape): void {
            $objWriter->writeAttribute('style:parent-style-name', 'standard');
            // style:graphic-properties
            $objWriter->startElement('style:graphic-properties');
            $objWriter->writeAttribute('style:mirror', 'none');
            $this->writeStylePartShadow($objWriter, $shape->getShadow());
            if (is_bool($shape->hasAutoShrinkVertical())) {
                $objWriter->writeAttribute('draw:auto-grow-height', var_export($shape->hasAutoShrinkVertical(), true));
            }
            if (is_bool($shape->hasAutoShrinkHorizontal())) {
                $objWriter->writeAttribute('draw:auto-grow-width', var_export($shape->hasAutoShrinkHorizontal(), true));
            }
            // Fill
            if (in_array($shape->getFill()->getFillType(), Fill::PATTERN_TYPES, true)) {
                $this->writePatternFill($objWriter, $shape->getFill());
            } else {
                switch ($shape->getFill()->getFillType()) {
                    case Fill::FILL_GRADIENT_LINEAR:
                    case Fill::FILL_GRADIENT_PATH:
                        $objWriter->writeAttribute('draw:fill', 'gradient');
                        $objWriter->writeAttribute('draw:fill-gradient-name', 'gradient_' . $shape->getFill()->getHashCode());

                        break;
                    case Fill::FILL_SOLID:
                        $objWriter->writeAttribute('draw:fill', 'solid');
                        $objWriter->writeAttribute('draw:fill-color', '#' . $shape->getFill()->getStartColor()->getRGB());

                        break;
                    case Fill::FILL_UNSET:
                        // Nobody named a fill, so the style names none and `standard` paints the shape

                        break;
                    case Fill::FILL_NONE:
                    default:
                        $objWriter->writeAttribute('draw:fill', 'none');
                        $objWriter->writeAttribute('draw:fill-color', '#' . $shape->getFill()->getStartColor()->getRGB());

                        break;
                }
            }
            // Border
            if (Border::LINE_NONE == $shape->getBorder()->getLineStyle()) {
                $objWriter->writeAttribute('draw:stroke', 'none');
            } else {
                // A border that names no colour is written without `svg:stroke-color`, which the
                // attribute being optional lets the parent style answer instead
                $borderColor = $shape->getBorder()->getColor();
                if (null !== $borderColor) {
                    $objWriter->writeAttribute('svg:stroke-color', '#' . $borderColor->getRGB());
                }
                $objWriter->writeAttribute('svg:stroke-width', number_format(CommonDrawing::pointsToCentimeters($shape->getBorder()->getLineWidth()), 3, '.', '') . 'cm');
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
                        $objWriter->writeAttribute('draw:stroke-dash', 'strokeDash_' . $shape->getBorder()->getDashStyle());

                        break;
                    default:
                        $objWriter->writeAttribute('draw:stroke', 'none');

                        break;
                }
            }

            $objWriter->writeAttribute('fo:wrap-option', 'wrap');
            // The writing mode of the frame orders its columns; the direction of the text comes from
            // the writing mode each paragraph style states for itself.
            $objWriter->writeAttribute('style:writing-mode', $shape->isColumnsRTL() ? 'rl-tb' : 'lr-tb');
            // style:graphic-properties > style:columns
            // An element rather than an attribute, so it is written after every attribute of the block
            if ($shape->getColumns() > 1) {
                $objWriter->startElement('style:columns');
                $objWriter->writeAttribute('fo:column-count', $shape->getColumns());
                $objWriter->writeAttribute(
                    'fo:column-gap',
                    number_format(CommonDrawing::pixelsToCentimeters($shape->getColumnSpacing()), 3, '.', '') . 'cm'
                );
                $objWriter->endElement();
            }
            // > style:graphic-properties
            $objWriter->endElement();
        }, $shape);

        $paragraphs = $shape->getParagraphs();
        $paragraphId = 0;
        foreach ($paragraphs as $paragraph) {
            ++$paragraphId;

            // Style des paragraphes
            $this->addParagraphStyle($paragraph);

            // Style des listes
            // Only a paragraph that asks for a marker is written inside a `text:list`, and only
            // that list names the style, so collecting one for any other paragraph puts a
            // `text:list-style` in the file that nothing points at. The condition is the one
            // writeShapeTxt() opens the list on.
            if ('' !== $this->getListStyleName($paragraph)) {
                $bulletStyleHashCode = $paragraph->getBulletStyle()->getHashCode();
                if (!isset($this->arrStyleBullet[$bulletStyleHashCode])) {
                    $this->arrStyleBullet[$bulletStyleHashCode]['oStyle'] = $paragraph->getBulletStyle();
                    $this->arrStyleBullet[$bulletStyleHashCode]['level'] = '';
                }
                if (false === strpos($this->arrStyleBullet[$bulletStyleHashCode]['level'], ';' . $paragraph->getAlignment()->getLevel())) {
                    $this->arrStyleBullet[$bulletStyleHashCode]['level'] .= ';' . $paragraph->getAlignment()->getLevel();
                    $this->arrStyleBullet[$bulletStyleHashCode]['oAlign_' . $paragraph->getAlignment()->getLevel()] = $paragraph->getAlignment();
                }
            }

            $richtexts = $paragraph->getRichTextElements();
            $richtextId = 0;
            foreach ($richtexts as $richtext) {
                ++$richtextId;
                // Not a line break
                if ($richtext instanceof Run) {
                    // Style des font text
                    $this->addTextStyle($richtext);
                }
            }
        }
    }

    /**
     * Name the automatic style an AbstractDrawingAdapter wears.
     */
    protected function addDrawingStyle(AbstractDrawingAdapter $shape): void
    {
        $this->shareAutomaticStyle('graphic', function (XMLWriter $objWriter) use ($shape): void {
            $objWriter->writeAttribute('style:parent-style-name', 'standard');

            // style:graphic-properties
            $objWriter->startElement('style:graphic-properties');
            $objWriter->writeAttribute('draw:stroke', 'none');
            $objWriter->writeAttribute('style:mirror', 'none');
            $this->writeStylePartFill($objWriter, $shape->getFill());
            $this->writeStylePartShadow($objWriter, $shape->getShadow());
            $objWriter->endElement();
        }, $shape);
    }

    /**
     * Name the automatic style a Line shape wears.
     */
    protected function addLineStyle(Line $shape): void
    {
        $this->shareAutomaticStyle('graphic', function (XMLWriter $objWriter) use ($shape): void {
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
            $borderColor = $shape->getBorder()->getColor();
            if (null !== $borderColor) {
                $objWriter->writeAttribute('svg:stroke-color', '#' . $borderColor->getRGB());
            }
            $objWriter->writeAttribute('svg:stroke-width', Text::numberFormat(CommonDrawing::pointsToCentimeters($shape->getBorder()->getLineWidth()), 3) . 'cm');
            $this->writeStylePartShadow($objWriter, $shape->getShadow());
            $objWriter->endElement();
        }, $shape);
    }

    /**
     * The value of an ODF fo:border property: a width, a style and a colour.
     */
    protected function getBorderValue(Border $border): string
    {
        $value = Text::numberFormat($border->getLineWidth() / 1.75, 2) . 'pt '
            . $this->getBorderStyle($border);

        // `fo:border` takes the CSS2 shorthand, where the colour is the optional third part: a
        // border that names none is written as a width and a style, and is painted in the colour
        // the cell's text is
        $color = $border->getColor();

        return null === $color ? $value : $value . ' #' . $color->getRGB();
    }

    /**
     * The style component of an ODF fo:border, which takes the CSS2 border styles. Those are fewer
     * than the OOXML line styles, so the compound lines all land on `double` and the ten dash
     * patterns on `dashed` or `dotted` -- the closest CSS can say.
     */
    protected function getBorderStyle(Border $border): string
    {
        if (Border::LINE_NONE === $border->getLineStyle()) {
            return 'none';
        }

        switch ($border->getDashStyle()) {
            case Border::DASH_SOLID:
                break;
            case Border::DASH_DOT:
            case Border::DASH_SYSDOT:
                return 'dotted';
            default:
                return 'dashed';
        }

        return Border::LINE_SINGLE === $border->getLineStyle() ? 'solid' : 'double';
    }

    /**
     * Name the automatic styles the rows and the cells of a Table shape wear.
     */
    protected function addTableStyle(Table $shape): void
    {
        foreach ($shape->getRows() as $shapeRow) {
            $this->shareAutomaticStyle('table-row', function (XMLWriter $objWriter) use ($shapeRow): void {
                // style:table-row-properties
                $objWriter->startElement('style:table-row-properties');
                $objWriter->writeAttribute('style:row-height', Text::numberFormat(CommonDrawing::pointsToCentimeters($shapeRow->getHeight()), 3) . 'cm');
                $objWriter->endElement();
            }, $shapeRow);

            foreach ($shapeRow->getCells() as $shapeCell) {
                // The body a cell writes is read out of two objects, which is why it is a closure
                // and not a string: the cell, and the row it sits in, because a cell that named no
                // fill of its own is painted with the row's
                $this->shareAutomaticStyle('table-cell', function (XMLWriter $objWriter) use ($shapeCell, $shapeRow): void {
                    // A cell that was given no fill of its own is painted with the fill of its row.
                    // A cell set to `FILL_NONE` asked for no fill and stays transparent, even where
                    // its row is painted.
                    $cellFill = $shapeCell->getFill();
                    if (Fill::FILL_UNSET == $cellFill->getFillType()) {
                        $cellFill = $shapeRow->getFill();
                    }

                    // Note : This element is not valid in the Schema 1.2
                    // style:graphic-properties
                    if (Fill::FILL_NONE != $cellFill->getFillType()
                        && Fill::FILL_UNSET != $cellFill->getFillType()
                    ) {
                        $objWriter->startElement('style:graphic-properties');
                        if (Fill::FILL_SOLID == $cellFill->getFillType()) {
                            $objWriter->writeAttribute('draw:fill', 'solid');
                            $objWriter->writeAttribute('draw:fill-color', '#' . $cellFill->getStartColor()->getRGB());
                        }
                        if (Fill::FILL_GRADIENT_LINEAR == $cellFill->getFillType()) {
                            $objWriter->writeAttribute('draw:fill', 'gradient');
                            $objWriter->writeAttribute('draw:fill-gradient-name', 'gradient_' . $cellFill->getHashCode());
                        }
                        if (in_array($cellFill->getFillType(), Fill::PATTERN_TYPES, true)) {
                            $this->writePatternFill($objWriter, $cellFill);
                        }
                        $objWriter->endElement();
                    }
                    // >style:graphic-properties

                    // style:paragraph-properties
                    $objWriter->startElement('style:paragraph-properties');
                    $cellBorders = $shapeCell->getBorders();
                    $cellBordersBottomHashCode = $cellBorders->getBottom()->getHashCode();
                    if ($cellBordersBottomHashCode == $cellBorders->getTop()->getHashCode()
                        && $cellBordersBottomHashCode == $cellBorders->getLeft()->getHashCode()
                        && $cellBordersBottomHashCode == $cellBorders->getRight()->getHashCode()) {
                        $objWriter->writeAttribute('fo:border', $this->getBorderValue($cellBorders->getBottom()));
                    } else {
                        $objWriter->writeAttribute('fo:border-bottom', $this->getBorderValue($cellBorders->getBottom()));
                        $objWriter->writeAttribute('fo:border-top', $this->getBorderValue($cellBorders->getTop()));
                        $objWriter->writeAttribute('fo:border-right', $this->getBorderValue($cellBorders->getRight()));
                        $objWriter->writeAttribute('fo:border-left', $this->getBorderValue($cellBorders->getLeft()));
                    }
                    $objWriter->endElement();
                    // >style:paragraph-properties
                }, $shapeCell);

                foreach ($shapeCell->getParagraphs() as $shapeParagraph) {
                    foreach ($shapeParagraph->getRichTextElements() as $shapeRichText) {
                        if ($shapeRichText instanceof Run) {
                            // Style des font text
                            $this->addTextStyle($shapeRichText);
                        }
                    }
                }
            }
        }
    }

    /**
     * Write the slide note.
     */
    protected function writeSlideNote(XMLWriter $objWriter, Note $note): void
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
     * Name the automatic style a slide wears.
     */
    protected function addSlideStyle(Slide $slide, int $incPage): void
    {
        $this->shareAutomaticStyle('drawing-page', function (XMLWriter $objWriter) use ($slide, $incPage): void {
            // style:style/style:drawing-page-properties
            $objWriter->startElement('style:drawing-page-properties');
            $objWriter->writeAttributeIf(!$slide->isVisible(), 'presentation:visibility', 'hidden');
            if (null !== ($oTransition = $slide->getTransition())) {
                $objWriter->writeAttribute('presentation:duration', 'PT' . number_format($oTransition->getAdvanceTimeTrigger() / 1000, 6, '.', '') . 'S');
                $objWriter->writeAttributeIf($oTransition->hasManualTrigger(), 'presentation:transition-type', 'manual');
                $objWriter->writeAttributeIf($oTransition->hasTimeTrigger(), 'presentation:transition-type', 'automatic');
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

                // http://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part1.html#property-presentation_transition-style
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
                    case Transition::TRANSITION_CIRCLE:
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
            $oBackground = $slide->getBackground();
            if ($oBackground instanceof Slide\AbstractBackground) {
                $objWriter->writeAttribute('presentation:background-visible', 'true');
                if ($oBackground instanceof Slide\Background\Color) {
                    $objWriter->writeAttribute('draw:fill', 'solid');
                    $objWriter->writeAttribute('draw:fill-color', '#' . $oBackground->getColor()->getRGB());
                }
                if ($oBackground instanceof Slide\Background\Image) {
                    $objWriter->writeAttribute('draw:fill', 'bitmap');
                    $objWriter->writeAttribute('draw:fill-image-name', 'background_' . $incPage);
                    $objWriter->writeAttribute('style:repeat', 'stretch');
                }
            }
            $objWriter->endElement();
        }, $slide);
    }

    protected function writeStylePartFill(XMLWriter $objWriter, Fill $oFill): void
    {
        switch ($oFill->getFillType()) {
            case Fill::FILL_UNSET:
                // Nobody named a fill, so the style names none and the parent style paints the shape

                break;
            case Fill::FILL_SOLID:
                $objWriter->writeAttribute('draw:fill', 'solid');
                $objWriter->writeAttribute('draw:fill-color', '#' . $oFill->getStartColor()->getRGB());

                break;
            case Fill::FILL_NONE:
            default:
                $objWriter->writeAttribute('draw:fill', 'none');

                break;
        }
    }

    /**
     * @todo Improve for supporting any direction (https://sinepost.wordpress.com/2012/02/16/theyve-got-atan-you-want-atan2/)
     */
    protected function writeStylePartShadow(XMLWriter $objWriter, Shadow $oShadow): void
    {
        if (!$oShadow->isVisible()) {
            return;
        }
        $objWriter->writeAttribute('draw:shadow', 'visible');
        $objWriter->writeAttribute('draw:shadow-color', '#' . $oShadow->getColor()->getRGB());

        $distanceCms = CommonDrawing::pixelsToCentimeters((int) $oShadow->getDistance());
        if (0 == $oShadow->getDirection() || 360 == $oShadow->getDirection()) {
            $objWriter->writeAttribute('draw:shadow-offset-x', $distanceCms . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '0cm');
        } elseif (45 == $oShadow->getDirection()) {
            $objWriter->writeAttribute('draw:shadow-offset-x', $distanceCms . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', $distanceCms . 'cm');
        } elseif (90 == $oShadow->getDirection()) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '0cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', $distanceCms . 'cm');
        } elseif (135 == $oShadow->getDirection()) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '-' . $distanceCms . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', $distanceCms . 'cm');
        } elseif (180 == $oShadow->getDirection()) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '-' . $distanceCms . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '0cm');
        } elseif (225 == $oShadow->getDirection()) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '-' . $distanceCms . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '-' . $distanceCms . 'cm');
        } elseif (270 == $oShadow->getDirection()) {
            $objWriter->writeAttribute('draw:shadow-offset-x', '0cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '-' . $distanceCms . 'cm');
        } elseif (315 == $oShadow->getDirection()) {
            $objWriter->writeAttribute('draw:shadow-offset-x', $distanceCms . 'cm');
            $objWriter->writeAttribute('draw:shadow-offset-y', '-' . $distanceCms . 'cm');
        }
        $objWriter->writeAttribute('draw:shadow-opacity', (100 - $oShadow->getAlpha()) . '%');
    }
}
