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
use PhpOffice\PhpPresentation\Style\Font;

abstract class AbstractDecoratorWriter extends \PhpOffice\PhpPresentation\Writer\AbstractDecoratorWriter
{
    /**
     * OOXML names an underline with a single token; ODF spells the same thing over a style, a type
     * and a width, and has no name of its own for a few of the eighteen.
     *
     * @var array<string, array{0: string, 1: string, 2: null|string}>
     */
    private const UNDERLINE_ODF = [
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
     * @var array<string, string>
     */
    private const STRIKETHROUGH_ODF = [
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
