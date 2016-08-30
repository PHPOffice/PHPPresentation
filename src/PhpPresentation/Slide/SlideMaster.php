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
namespace PhpOffice\PhpPresentation\Slide;

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\ShapeContainerInterface;
use PhpOffice\PhpPresentation\Slide\Background\Color as BackgroundColor;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\ColorMap;
use PhpOffice\PhpPresentation\Style\SchemeColor;
use PhpOffice\PhpPresentation\Style\TextStyle;

/**
 * Class SlideMaster
 */
class SlideMaster extends AbstractSlide implements ComparableInterface, ShapeContainerInterface
{
    /**
     * Collection of Slide objects
     *
     * @var \PhpOffice\PhpPresentation\Slide\SlideLayout[]
     */
    protected $slideLayouts = array();
    /**
     * Mapping of colors to the theme
     *
     * @var \PhpOffice\PhpPresentation\Style\ColorMap
     */
    public $colorMap;
    /**
     * @var \PhpOffice\PhpPresentation\Style\TextStyle
     */
    protected $textStyles;
    /**
     * @var \PhpOffice\PhpPresentation\Style\SchemeColor[]
     */
    protected $arraySchemeColor = array();
    /**
     * @var array
     */
    protected $defaultSchemeColor = array(
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
    );

    /**
     * Create a new slideMaster
     *
     * @param PhpPresentation $pParent
     */
    public function __construct(PhpPresentation $pParent = null)
    {
        // Set parent
        $this->parent = $pParent;
        // Shape collection
        $this->shapeCollection = new \ArrayObject();
        // Set identifier
        $this->identifier = md5(rand(0, 9999) . time());
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
     * Create a slideLayout and add it to this presentation
     *
     * @return \PhpOffice\PhpPresentation\Slide\SlideLayout
     */
    public function createSlideLayout()
    {
        $newSlideLayout = new SlideLayout($this);
        $this->addSlideLayout($newSlideLayout);
        return $newSlideLayout;
    }

    /**
     * Add slideLayout
     *
     * @param  \PhpOffice\PhpPresentation\Slide\SlideLayout $slideLayout
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Slide\SlideLayout
     */
    public function addSlideLayout(SlideLayout $slideLayout = null)
    {
        $this->slideLayouts[] = $slideLayout;
        return $slideLayout;
    }

    /**
     * @return SlideLayout[]
     */
    public function getAllSlideLayouts()
    {
        return $this->slideLayouts;
    }

    /**
     * @return TextStyle
     */
    public function getTextStyles()
    {
        return $this->textStyles;
    }

    /**
     * @param TextStyle $textStyle
     * @return $this
     */
    public function setTextStyles(TextStyle $textStyle)
    {
        $this->textStyles = $textStyle;
        return $this;
    }

    /**
     * @param SchemeColor $schemeColor
     * @return $this
     */
    public function addSchemeColor(SchemeColor $schemeColor)
    {
        $this->arraySchemeColor[$schemeColor->getValue()] = $schemeColor;
        return $this;
    }

    /**
     * @return \PhpOffice\PhpPresentation\Style\SchemeColor[]
     */
    public function getAllSchemeColors()
    {
        return $this->arraySchemeColor;
    }
}
