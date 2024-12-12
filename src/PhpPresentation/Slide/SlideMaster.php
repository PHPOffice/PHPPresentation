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

namespace PhpOffice\PhpPresentation\Slide;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Slide\Background\Color as BackgroundColor;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\ColorMap;
use PhpOffice\PhpPresentation\Style\SchemeColor;
use PhpOffice\PhpPresentation\Style\TextStyle;

class SlideMaster extends AbstractSlide implements ComparableInterface, ShapeContainerInterface
{
    /**
     * Collection of Slide objects.
     *
     * @var array<SlideLayout>
     */
    protected $slideLayouts = [];

    /**
     * Mapping of colors to the theme.
     *
     * @var ColorMap
     */
    public $colorMap;

    /**
     * @var TextStyle
     */
    protected $textStyles;

    /**
     * @var array<SchemeColor>
     */
    protected $arraySchemeColor = [];

    /**
     * @var array<string, string>
     */
    protected $defaultSchemeColor = [
        'dk1' => '000000',
        'lt1' => 'FFFFFF',
        'dk2' => '1F497D',
        'lt2' => 'EEECE1',
        'accent1' => '4F81BD',
        'accent2' => 'C0504D',
        'accent3' => '9BBB59',
        'accent4' => '8064A2',
        'accent5' => '4BACC6',
        'accent6' => 'F79646',
        'hlink' => '0000FF',
        'folHlink' => '800080',
    ];

    /**
     * Create a new slideMaster.
     */
    public function __construct(?PhpPresentation $pParent = null)
    {
        // Set parent
        $this->parent = $pParent;
        // Set identifier
        $this->identifier = md5(mt_rand(0, 9999) . time());
        // Set a basic colorMap
        $this->colorMap = new ColorMap();
        // Set a white background
        $this->background = new BackgroundColor();
        $this->background->setColor(new Color(Color::COLOR_WHITE));
        // Set basic textStyles
        $this->textStyles = new TextStyle(true);
        // Set basic scheme colors
        foreach ($this->defaultSchemeColor as $key => $value) {
            $oSchemeColor = new SchemeColor();
            $oSchemeColor->setValue($key);
            $oSchemeColor->setRGB($value);
            $this->addSchemeColor($oSchemeColor);
        }
    }

    /**
     * Create a slideLayout and add it to this presentation.
     */
    public function createSlideLayout(): SlideLayout
    {
        $newSlideLayout = new SlideLayout($this);
        $this->addSlideLayout($newSlideLayout);

        return $newSlideLayout;
    }

    /**
     * Add slideLayout.
     */
    public function addSlideLayout(?SlideLayout $slideLayout = null): SlideLayout
    {
        $this->slideLayouts[] = $slideLayout;

        return $slideLayout;
    }

    /**
     * @return array<SlideLayout>
     */
    public function getAllSlideLayouts(): array
    {
        return $this->slideLayouts;
    }

    public function getTextStyles(): TextStyle
    {
        return $this->textStyles;
    }

    public function setTextStyles(TextStyle $textStyle): self
    {
        $this->textStyles = $textStyle;

        return $this;
    }

    public function addSchemeColor(SchemeColor $schemeColor): self
    {
        $this->arraySchemeColor[$schemeColor->getValue()] = $schemeColor;

        return $this;
    }

    /**
     * @return array<SchemeColor>
     */
    public function getAllSchemeColors(): array
    {
        return $this->arraySchemeColor;
    }
}
