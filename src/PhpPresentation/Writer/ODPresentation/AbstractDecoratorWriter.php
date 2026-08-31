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

use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Slide\AbstractBackground;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;

abstract class AbstractDecoratorWriter extends \PhpOffice\PhpPresentation\Writer\AbstractDecoratorWriter
{
    /**
     * OOXML names an underline with a single token; ODF spells the same thing over a style, a type
     * and a width, and has no name of its own for a few of the eighteen.
     *
     * @var array<string, array{0: string, 1: string, 2: null|string}>
     */
    protected const UNDERLINE_ODF = [
        Font::UNDERLINE_DASH => ['dash', 'single', null],
        Font::UNDERLINE_DASHHEAVY => ['dash', 'single', 'bold'],
        Font::UNDERLINE_DASHLONG => ['long-dash', 'single', null],
        Font::UNDERLINE_DASHLONGHEAVY => ['long-dash', 'single', 'bold'],
        Font::UNDERLINE_DOTHASH => ['dot-dash', 'single', null],
        Font::UNDERLINE_DOTHASHHEAVY => ['dot-dash', 'single', 'bold'],
        Font::UNDERLINE_DOTDOTDASH => ['dot-dot-dash', 'single', null],
        Font::UNDERLINE_DOTDOTDASHHEAVY => ['dot-dot-dash', 'single', 'bold'],
        Font::UNDERLINE_DOTTED => ['dotted', 'single', null],
        Font::UNDERLINE_DOTTEDHEAVY => ['dotted', 'single', 'bold'],
        Font::UNDERLINE_DOUBLE => ['solid', 'double', null],
        Font::UNDERLINE_HEAVY => ['solid', 'single', 'bold'],
        Font::UNDERLINE_SINGLE => ['solid', 'single', null],
        Font::UNDERLINE_WAVY => ['wave', 'single', null],
        Font::UNDERLINE_WAVYDOUBLE => ['wave', 'double', null],
        Font::UNDERLINE_WAVYHEAVY => ['wave', 'single', 'bold'],
        Font::UNDERLINE_WORDS => ['solid', 'single', null],
    ];

    /**
     * DrawingML names fifty-four preset patterns; ODF has `draw:hatch`, which is one, two or three
     * families of parallel lines and nothing else.
     *
     * The patterns below are the ones that are lines, mapped to the style, the angle and the
     * spacing that draw them. The other twenty-two -- the bricks, the confetti, the diamonds, the
     * sphere, the weave of a plaid, the percentage screens -- are not lines at all, and ODF has no
     * way to say them. Those are painted solid in the colour of the pattern rather than dropped:
     * a shape filled with the wrong texture is nearer what was asked for than an empty one.
     *
     * @var array<string, array{0: string, 1: int, 2: string}>
     */
    protected const HATCH_ODF = [
        Fill::FILL_PATTERN_HORZ => ['single', 0, '0.1cm'],
        Fill::FILL_PATTERN_VERT => ['single', 90, '0.1cm'],
        Fill::FILL_PATTERN_LTHORZ => ['single', 0, '0.2cm'],
        Fill::FILL_PATTERN_LTVERT => ['single', 90, '0.2cm'],
        Fill::FILL_PATTERN_DKHORZ => ['single', 0, '0.05cm'],
        Fill::FILL_PATTERN_DKVERT => ['single', 90, '0.05cm'],
        Fill::FILL_PATTERN_NARHORZ => ['single', 0, '0.05cm'],
        Fill::FILL_PATTERN_NARVERT => ['single', 90, '0.05cm'],
        Fill::FILL_PATTERN_DASHHORZ => ['single', 0, '0.2cm'],
        Fill::FILL_PATTERN_DASHVERT => ['single', 90, '0.2cm'],
        Fill::FILL_PATTERN_CROSS => ['double', 0, '0.1cm'],
        Fill::FILL_PATTERN_DNDIAG => ['single', 315, '0.1cm'],
        Fill::FILL_PATTERN_UPDIAG => ['single', 45, '0.1cm'],
        Fill::FILL_PATTERN_LTDNDIAG => ['single', 315, '0.2cm'],
        Fill::FILL_PATTERN_LTUPDIAG => ['single', 45, '0.2cm'],
        Fill::FILL_PATTERN_DKDNDIAG => ['single', 315, '0.05cm'],
        Fill::FILL_PATTERN_DKUPDIAG => ['single', 45, '0.05cm'],
        Fill::FILL_PATTERN_WDDNDIAG => ['single', 315, '0.2cm'],
        Fill::FILL_PATTERN_WDUPDIAG => ['single', 45, '0.2cm'],
        Fill::FILL_PATTERN_DASHDNDIAG => ['single', 315, '0.2cm'],
        Fill::FILL_PATTERN_DASHUPDIAG => ['single', 45, '0.2cm'],
        Fill::FILL_PATTERN_DIAGCROSS => ['double', 45, '0.1cm'],
        Fill::FILL_PATTERN_SMGRID => ['double', 0, '0.05cm'],
        Fill::FILL_PATTERN_LGGRID => ['double', 0, '0.2cm'],
        Fill::FILL_PATTERN_DOTGRID => ['double', 0, '0.2cm'],
        Fill::FILL_PATTERN_SMCHECK => ['double', 45, '0.05cm'],
        Fill::FILL_PATTERN_LGCHECK => ['double', 45, '0.2cm'],
        Fill::FILL_PATTERN_TRELLIS => ['double', 45, '0.1cm'],
        Fill::FILL_PATTERN_WEAVE => ['double', 45, '0.1cm'],
        Fill::FILL_PATTERN_PLAID => ['triple', 0, '0.1cm'],
        Fill::FILL_PATTERN_ZIGZAG => ['single', 0, '0.1cm'],
        Fill::FILL_PATTERN_WAVE => ['single', 0, '0.2cm'],
    ];

    /**
     * @var array<string, string>
     */
    protected const STRIKETHROUGH_ODF = [
        Font::STRIKE_SINGLE => 'single',
        Font::STRIKE_DOUBLE => 'double',
    ];

    /**
     * The underline and the strikethrough of a font, written into the `style:text-properties` the
     * caller has open. Both are one attribute family for the whole run, unlike the family, the
     * size and the weight, which ODF spells once per script.
     */
    protected function writeFontStates(XMLWriter $objWriter, Font $font): void
    {
        if (isset(self::UNDERLINE_ODF[$font->getUnderline()])) {
            [$style, $type, $width] = self::UNDERLINE_ODF[$font->getUnderline()];
            $objWriter->writeAttribute('style:text-underline-style', $style);
            $objWriter->writeAttribute('style:text-underline-type', $type);
            if (null !== $width) {
                $objWriter->writeAttribute('style:text-underline-width', $width);
            }
            if (Font::UNDERLINE_WORDS === $font->getUnderline()) {
                $objWriter->writeAttribute('style:text-underline-mode', 'skip-white-space');
            }
        }
        if (isset(self::STRIKETHROUGH_ODF[$font->getStrikethrough()])) {
            $objWriter->writeAttribute('style:text-line-through-style', 'solid');
            $objWriter->writeAttribute('style:text-line-through-type', self::STRIKETHROUGH_ODF[$font->getStrikethrough()]);
        }
    }

    /**
     * The fill attributes of a pattern, into the `style:graphic-properties` the caller has open.
     *
     * A pattern ODF can draw is a hatch, whose lines are the start colour and whose ground is the
     * end colour; `draw:fill-hatch-solid` is what says the ground is painted at all, and without it
     * the lines are drawn over whatever is behind the shape. A pattern ODF cannot draw is painted
     * in the colour of its lines.
     */
    protected function writePatternFill(XMLWriter $objWriter, Fill $fill): void
    {
        if (!isset(self::HATCH_ODF[$fill->getFillType()])) {
            $objWriter->writeAttribute('draw:fill', 'solid');
            $objWriter->writeAttribute('draw:fill-color', '#' . $fill->getStartColor()->getRGB());

            return;
        }

        $objWriter->writeAttribute('draw:fill', 'hatch');
        $objWriter->writeAttribute('draw:fill-hatch-name', 'hatch_' . $fill->getHashCode());
        $objWriter->writeAttribute('draw:fill-hatch-solid', 'true');
        $objWriter->writeAttribute('draw:fill-color', '#' . $fill->getEndColor()->getRGB());
    }

    /**
     * @var Chart[]
     */
    protected $arrayChart;

    /**
     * The background of the master page, which every slide is drawn on top of.
     *
     * The Writer has one master page, `Standard`, so the first master slide is the one that owns it.
     */
    protected function getMasterBackground(): ?AbstractBackground
    {
        foreach ($this->getPresentation()->getAllMasterSlides() as $oMasterSlide) {
            return $oMasterSlide->getBackground();
        }

        return null;
    }

    /**
     * @return Chart[]
     */
    public function getArrayChart()
    {
        return $this->arrayChart;
    }

    /**
     * @param Chart[] $arrayChart
     *
     * @return AbstractDecoratorWriter
     */
    public function setArrayChart($arrayChart)
    {
        $this->arrayChart = $arrayChart;

        return $this;
    }
}
