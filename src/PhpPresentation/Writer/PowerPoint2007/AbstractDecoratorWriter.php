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

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use DK\OpenXml\OpenXmlPackage;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;

abstract class AbstractDecoratorWriter extends \PhpOffice\PhpPresentation\Writer\AbstractDecoratorWriter
{
    /**
     * What a media part holds, by the extension it is written under.
     *
     * These are the types the Writer produces itself and the ones a package declares once for the
     * extension rather than once for each part; anything else a drawing carries is written with
     * the type the drawing reports.
     *
     * @var array<string, string>
     */
    public const MEDIA_CONTENT_TYPES = [
        'gif' => 'image/gif',
        'jpeg' => 'image/jpeg',
        'jpg' => 'image/jpeg',
        'png' => 'image/png',
        'svg' => 'image/svg+xml',
        'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ];

    abstract public function render(): OpenXmlPackage;

    /**
     * The type to write a media part with.
     *
     * The map is preferred over what the drawing reports, because an environment can report
     * `image/svg` for an SVG, which is not the registered type.
     */
    protected function mediaContentType(string $extension, ?string $reported = null): string
    {
        return self::MEDIA_CONTENT_TYPES[strtolower($extension)]
            ?? ($reported ?? 'application/octet-stream');
    }

    /**
     * @var OpenXmlPackage
     */
    protected $oPackage;

    /**
     * @return $this
     */
    public function setPackage(OpenXmlPackage $oPackage)
    {
        $this->oPackage = $oPackage;

        return $this;
    }

    public function getPackage(): OpenXmlPackage
    {
        return $this->oPackage;
    }

    /**
     * Say that one part points at another.
     *
     * The package writes the `.rels` parts itself, so a relationship is declared rather than
     * spelled out. The identifier stays the caller's to choose: the parts that hold a reference
     * name it -- `r:embed="rId2"` -- and are written by their own decorator, which counts the
     * relationships in the same order this is called in.
     *
     * @param null|string $source Part the relationship belongs to, null for the package itself
     * @param int $pId Relationship ID. rId will be prepended!
     * @param string $pType Relationship type
     * @param string $pTarget Relationship target
     * @param string $pTargetMode Relationship target mode
     */
    protected function writeRelationship(?string $source, int $pId, string $pType, string $pTarget, string $pTargetMode = ''): void
    {
        $this->oPackage->addRelationship($pType, $pTarget, 'External' === $pTargetMode, 'rId' . $pId, $source);
    }

    /**
     * Write Border.
     *
     * @param XMLWriter $objWriter XML Writer
     * @param Border $pBorder Border
     * @param string $pElementName Element name
     */
    protected function writeBorder(XMLWriter $objWriter, Border $pBorder, string $pElementName = 'L', bool $isMarker = false): void
    {
        if (Border::LINE_NONE == $pBorder->getLineStyle() && '' == $pElementName && !$isMarker) {
            return;
        }

        // Line style
        $lineStyle = $pBorder->getLineStyle();
        if (Border::LINE_NONE == $lineStyle) {
            $lineStyle = Border::LINE_SINGLE;
        }

        // a:ln $pElementName
        $objWriter->startElement('a:ln' . $pElementName);
        $objWriter->writeAttribute('w', (int) CommonDrawing::pixelsToEmu($pBorder->getLineWidth()));
        $objWriter->writeAttribute('cap', 'flat');
        $objWriter->writeAttribute('cmpd', $lineStyle);
        $objWriter->writeAttribute('algn', 'ctr');

        // Fill?
        $borderColor = $pBorder->getColor();
        if (Border::LINE_NONE == $pBorder->getLineStyle()) {
            // a:noFill
            $objWriter->writeElement('a:noFill', null);
        } elseif (null !== $borderColor) {
            // a:solidFill
            $objWriter->startElement('a:solidFill');
            $this->writeColor($objWriter, $borderColor);
            $objWriter->endElement();
        }
        // A border that names no colour is written without a fill of its own. The group is optional
        // in CT_LineProperties, and leaving it out is how the line takes its colour from the theme;
        // the width, the compound and the dash below are written either way.

        // Dash
        $objWriter->startElement('a:prstDash');
        $objWriter->writeAttribute('val', $pBorder->getDashStyle());
        $objWriter->endElement();

        // a:round
        $objWriter->writeElement('a:round', null);

        // a:headEnd
        $objWriter->startElement('a:headEnd');
        $objWriter->writeAttribute('type', 'none');
        $objWriter->writeAttribute('w', 'med');
        $objWriter->writeAttribute('len', 'med');
        $objWriter->endElement();

        // a:tailEnd
        $objWriter->startElement('a:tailEnd');
        $objWriter->writeAttribute('type', 'none');
        $objWriter->writeAttribute('w', 'med');
        $objWriter->writeAttribute('len', 'med');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    protected function writeColor(XMLWriter $objWriter, Color $color, ?int $alpha = null): void
    {
        if (null === $alpha) {
            $alpha = $color->getAlpha();
        }

        // a:srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $color->getRGB());

        // a:alpha
        $objWriter->startElement('a:alpha');
        $objWriter->writeAttribute('val', $alpha * 1000);
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * @param null|Fill $pFill a data point that has no fill in the sparse map of its series arrives
     *                         here as `null`; a shape whose fill was taken away arrives as `FILL_UNSET`
     */
    protected function writeFill(XMLWriter $objWriter, ?Fill $pFill): void
    {
        // A fill nobody asked for is written as nothing at all, so that the shape takes the one its
        // theme or its placeholder gives it. This is what a `null` fill has always been written as.
        if (null === $pFill || Fill::FILL_UNSET == $pFill->getFillType()) {
            return;
        }

        // A fill refused outright is `a:noFill` -- without this it would fall through to the
        // pattern fill at the end
        if (Fill::FILL_NONE == $pFill->getFillType()) {
            $objWriter->writeElement('a:noFill');

            return;
        }

        // Is it a solid fill?
        if (Fill::FILL_SOLID == $pFill->getFillType()) {
            $this->writeSolidFill($objWriter, $pFill);

            return;
        }

        // Is it a gradient fill?
        if (Fill::FILL_GRADIENT_LINEAR == $pFill->getFillType() || Fill::FILL_GRADIENT_PATH == $pFill->getFillType()) {
            $this->writeGradientFill($objWriter, $pFill);

            return;
        }

        // Is it a pattern fill?
        $this->writePatternFill($objWriter, $pFill);
    }

    protected function writeSolidFill(XMLWriter $objWriter, Fill $pFill): void
    {
        // a:gradFill
        $objWriter->startElement('a:solidFill');
        $this->writeColor($objWriter, $pFill->getStartColor());
        $objWriter->endElement();
    }

    protected function writeGradientFill(XMLWriter $objWriter, Fill $pFill): void
    {
        // a:gradFill
        $objWriter->startElement('a:gradFill');

        // a:gsLst
        $objWriter->startElement('a:gsLst');
        // a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '0');
        $this->writeColor($objWriter, $pFill->getStartColor());
        $objWriter->endElement();

        // a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100000');
        $this->writeColor($objWriter, $pFill->getEndColor());
        $objWriter->endElement();

        $objWriter->endElement();

        // a:lin
        $objWriter->startElement('a:lin');
        $objWriter->writeAttribute('ang', CommonDrawing::degreesToAngle((int) $pFill->getRotation()));
        $objWriter->writeAttribute('scaled', '0');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    protected function writePatternFill(XMLWriter $objWriter, Fill $pFill): void
    {
        // a:pattFill
        $objWriter->startElement('a:pattFill');

        // The pattern itself. Without `prst` the element carries two colours and no pattern, which
        // is what every pattern fill written here used to come out as -- and it is valid, because
        // the schema has the attribute optional, so nothing ever said so. A fill type that names no
        // preset is written as it was before, rather than as an attribute the schema would refuse.
        if (in_array($pFill->getFillType(), Fill::PATTERN_TYPES, true)) {
            $objWriter->writeAttribute('prst', $pFill->getFillType());
        }

        // fgClr
        $objWriter->startElement('a:fgClr');

        $this->writeColor($objWriter, $pFill->getStartColor());

        $objWriter->endElement();

        // bgClr
        $objWriter->startElement('a:bgClr');

        $this->writeColor($objWriter, $pFill->getEndColor());

        $objWriter->endElement();

        $objWriter->endElement();
    }

    protected function writeOutline(XMLWriter $objWriter, ?Outline $oOutline): void
    {
        if (!$oOutline) {
            return;
        }
        // Width : pixels
        $width = CommonDrawing::pixelsToEmu($oOutline->getWidth());

        // a:ln
        $objWriter->startElement('a:ln');
        $objWriter->writeAttribute('w', $width);

        // Fill
        $this->writeFill($objWriter, $oOutline->getFill());

        // > a:ln
        $objWriter->endElement();
    }

    /**
     * Determine absolute zip path.
     */
    protected function absoluteZipPath(string $path): string
    {
        $path = str_replace([
            '/',
            '\\',
        ], DIRECTORY_SEPARATOR, $path);
        $parts = array_filter(explode(DIRECTORY_SEPARATOR, $path), function (string $var) {
            return (bool) strlen($var);
        });
        $absolutes = [];
        foreach ($parts as $part) {
            if ('.' == $part) {
                continue;
            }
            if ('..' == $part) {
                array_pop($absolutes);
            } else {
                $absolutes[] = $part;
            }
        }

        return implode('/', $absolutes);
    }
}
