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

use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;

abstract class AbstractDecoratorWriter extends \PhpOffice\PhpPresentation\Writer\AbstractDecoratorWriter
{
    /**
     * Write relationship.
     *
     * @param XMLWriter $objWriter XML Writer
     * @param int $pId Relationship ID. rId will be prepended!
     * @param string $pType Relationship type
     * @param string $pTarget Relationship target
     * @param string $pTargetMode Relationship target mode
     */
    protected function writeRelationship(XMLWriter $objWriter, int $pId, string $pType, string $pTarget, string $pTargetMode = ''): void
    {
        // Write relationship
        $objWriter->startElement('Relationship');
        $objWriter->writeAttribute('Id', 'rId' . (string) $pId);
        $objWriter->writeAttribute('Type', $pType);
        $objWriter->writeAttribute('Target', $pTarget);

        if ('' != $pTargetMode) {
            $objWriter->writeAttribute('TargetMode', $pTargetMode);
        }

        $objWriter->endElement();
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
        if (!($pBorder instanceof Border)) {
            return;
        }

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
        if (Border::LINE_NONE == $pBorder->getLineStyle()) {
            // a:noFill
            $objWriter->writeElement('a:noFill', null);
        } else {
            // a:solidFill
            $objWriter->startElement('a:solidFill');
            $this->writeColor($objWriter, $pBorder->getColor());
            $objWriter->endElement();
        }

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

    protected function writeFill(XMLWriter $objWriter, ?Fill $pFill): void
    {
        if (!$pFill) {
            return;
        }

        // Is it a fill?
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
