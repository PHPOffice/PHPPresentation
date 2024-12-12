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

namespace PhpOffice\PhpPresentation\Shape;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Style\Outline;

/**
 * AutoShape shape.
 *
 * @see : https://github.com/scanny/python-pptx/blob/eaa1e0fd3db28b03a353e116a5c7d2084dd87c26/pptx/enum/shapes.py
 */
class AutoShape extends AbstractShape implements ComparableInterface
{
    public const TYPE_10_POINT_STAR = 'star10';
    public const TYPE_12_POINT_STAR = 'star12';
    public const TYPE_16_POINT_STAR = 'star16';
    public const TYPE_24_POINT_STAR = 'star24';
    public const TYPE_32_POINT_STAR = 'star32';
    public const TYPE_4_POINT_STAR = 'star4';
    public const TYPE_5_POINT_STAR = 'star5';
    public const TYPE_6_POINT_STAR = 'star6';
    public const TYPE_7_POINT_STAR = 'star7';
    public const TYPE_8_POINT_STAR = 'star8';
    public const TYPE_ACTION_BUTTON_BACK_OR_PREVIOUS = 'actionButtonBackPrevious';
    public const TYPE_ACTION_BUTTON_BEGINNING = 'actionButtonBeginning';
    public const TYPE_ACTION_BUTTON_CUSTOM = 'actionButtonBlank';
    public const TYPE_ACTION_BUTTON_DOCUMENT = 'actionButtonDocument';
    public const TYPE_ACTION_BUTTON_END = 'actionButtonEnd';
    public const TYPE_ACTION_BUTTON_FORWARD_OR_NEXT = 'actionButtonForwardNext';
    public const TYPE_ACTION_BUTTON_HELP = 'actionButtonHelp';
    public const TYPE_ACTION_BUTTON_HOME = 'actionButtonHome';
    public const TYPE_ACTION_BUTTON_INFORMATION = 'actionButtonInformation';
    public const TYPE_ACTION_BUTTON_MOVIE = 'actionButtonMovie';
    public const TYPE_ACTION_BUTTON_RETURN = 'actionButtonReturn';
    public const TYPE_ACTION_BUTTON_SOUND = 'actionButtonSound';
    public const TYPE_ARC = 'arc';
    public const TYPE_BALLOON = 'wedgeRoundRectCallout';
    public const TYPE_BENT_ARROW = 'bentArrow';
    public const TYPE_BENT_UP_ARROW = 'bentUpArrow';
    public const TYPE_BEVEL = 'bevel';
    public const TYPE_BLOCK_ARC = 'blockArc';
    public const TYPE_CAN = 'can';
    public const TYPE_CHART_PLUS = 'chartPlus';
    public const TYPE_CHART_STAR = 'chartStar';
    public const TYPE_CHARTX = 'chartX';
    public const TYPE_CHEVRON = 'chevron';
    public const TYPE_CHORD = 'chord';
    public const TYPE_CIRCULAR_ARROW = 'circularArrow';
    public const TYPE_CLOUD = 'cloud';
    public const TYPE_CLOUD_CALLOUT = 'cloudCallout';
    public const TYPE_CORNER = 'corner';
    public const TYPE_CORNER_TABS = 'cornerTabs';
    public const TYPE_CROSS = 'plus';
    public const TYPE_CUBE = 'cube';
    public const TYPE_CURVED_DOWN_ARROW = 'curvedDownArrow';
    public const TYPE_CURVED_DOWN_RIBBON = 'ellipseRibbon';
    public const TYPE_CURVED_LEFT_ARROW = 'curvedLeftArrow';
    public const TYPE_CURVED_RIGHT_ARROW = 'curvedRightArrow';
    public const TYPE_CURVED_UP_ARROW = 'curvedUpArrow';
    public const TYPE_CURVED_UP_RIBBON = 'ellipseRibbon2';
    public const TYPE_DECAGON = 'decagon';
    public const TYPE_DIAGONALSTRIPE = 'diagStripe';
    public const TYPE_DIAMOND = 'diamond';
    public const TYPE_DODECAGON = 'dodecagon';
    public const TYPE_DONUT = 'donut';
    public const TYPE_DOUBLE_BRACE = 'bracePair';
    public const TYPE_DOUBLE_BRACKET = 'bracketPair';
    public const TYPE_DOUBLE_WAVE = 'doubleWave';
    public const TYPE_DOWN_ARROW = 'downArrow';
    public const TYPE_DOWN_ARROWCALLOUT = 'downArrowCallout';
    public const TYPE_DOWN_RIBBON = 'ribbon';
    public const TYPE_EXPLOSIONEXPLOSION1 = 'irregularSeal1';
    public const TYPE_EXPLOSIONEXPLOSION2 = 'irregularSeal2';
    public const TYPE_FLOWCHART_ALTERNATEPROCESS = 'flowChartAlternateProcess';
    public const TYPE_FLOWCHART_CARD = 'flowChartPunchedCard';
    public const TYPE_FLOWCHART_COLLATE = 'flowChartCollate';
    public const TYPE_FLOWCHART_CONNECTOR = 'flowChartConnector';
    public const TYPE_FLOWCHART_DATA = 'flowChartInputOutput';
    public const TYPE_FLOWCHART_DECISION = 'flowChartDecision';
    public const TYPE_FLOWCHART_DELAY = 'flowChartDelay';
    public const TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE = 'flowChartMagneticDrum';
    public const TYPE_FLOWCHART_DISPLAY = 'flowChartDisplay';
    public const TYPE_FLOWCHART_DOCUMENT = 'flowChartDocument';
    public const TYPE_FLOWCHART_EXTRACT = 'flowChartExtract';
    public const TYPE_FLOWCHART_INTERNAL_STORAGE = 'flowChartInternalStorage';
    public const TYPE_FLOWCHART_MAGNETIC_DISK = 'flowChartMagneticDisk';
    public const TYPE_FLOWCHART_MANUAL_INPUT = 'flowChartManualInput';
    public const TYPE_FLOWCHART_MANUAL_OPERATION = 'flowChartManualOperation';
    public const TYPE_FLOWCHART_MERGE = 'flowChartMerge';
    public const TYPE_FLOWCHART_MULTIDOCUMENT = 'flowChartMultidocument';
    public const TYPE_FLOWCHART_OFFLINE_STORAGE = 'flowChartOfflineStorage';
    public const TYPE_FLOWCHART_OFFPAGE_CONNECTOR = 'flowChartOffpageConnector';
    public const TYPE_FLOWCHART_OR = 'flowChartOr';
    public const TYPE_FLOWCHART_PREDEFINED_PROCESS = 'flowChartPredefinedProcess';
    public const TYPE_FLOWCHART_PREPARATION = 'flowChartPreparation';
    public const TYPE_FLOWCHART_PROCESS = 'flowChartProcess';
    public const TYPE_FLOWCHART_PUNCHEDTAPE = 'flowChartPunchedTape';
    public const TYPE_FLOWCHART_SEQUENTIAL_ACCESS_STORAGE = 'flowChartMagneticTape';
    public const TYPE_FLOWCHART_SORT = 'flowChartSort';
    public const TYPE_FLOWCHART_STORED_DATA = 'flowChartOnlineStorage';
    public const TYPE_FLOWCHART_SUMMING_JUNCTION = 'flowChartSummingJunction';
    public const TYPE_FLOWCHART_TERMINATOR = 'flowChartTerminator';
    public const TYPE_FOLDED_CORNER = 'foldedCorner';
    public const TYPE_FRAME = 'frame';
    public const TYPE_FUNNEL = 'funnel';
    public const TYPE_GEAR_6 = 'gear6';
    public const TYPE_GEAR_9 = 'gear9';
    public const TYPE_HALF_FRAME = 'halfFrame';
    public const TYPE_HEART = 'heart';
    public const TYPE_HEPTAGON = 'heptagon';
    public const TYPE_HEXAGON = 'hexagon';
    public const TYPE_HORIZONTAL_SCROLL = 'horizontalScroll';
    public const TYPE_ISOSCELES_TRIANGLE = 'triangle';
    public const TYPE_LEFT_ARROW = 'leftArrow';
    public const TYPE_LEFT_ARROW_CALLOUT = 'leftArrowCallout';
    public const TYPE_LEFT_BRACE = 'leftBrace';
    public const TYPE_LEFT_BRACKET = 'leftBracket';
    public const TYPE_LEFT_CIRCULAR_ARROW = 'leftCircularArrow';
    public const TYPE_LEFT_RIGHT_ARROW = 'leftRightArrow';
    public const TYPE_LEFT_RIGHT_ARROW_CALLOUT = 'leftRightArrowCallout';
    public const TYPE_LEFT_RIGHT_CIRCULAR_ARROW = 'leftRightCircularArrow';
    public const TYPE_LEFT_RIGHT_RIBBON = 'leftRightRibbon';
    public const TYPE_LEFT_RIGHT_UP_ARROW = 'leftRightUpArrow';
    public const TYPE_LEFT_UP_ARROW = 'leftUpArrow';
    public const TYPE_LIGHTNING_BOLT = 'lightningBolt';
    public const TYPE_LINE_CALLOUT_1 = 'borderCallout1';
    public const TYPE_LINE_CALLOUT_1_ACCENT_BAR = 'accentCallout1';
    public const TYPE_LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR = 'accentBorderCallout1';
    public const TYPE_LINE_CALLOUT_1_NO_BORDER = 'callout1';
    public const TYPE_LINE_CALLOUT_2 = 'borderCallout2';
    public const TYPE_LINE_CALLOUT_2_ACCENT_BAR = 'accentCallout2';
    public const TYPE_LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR = 'accentBorderCallout2';
    public const TYPE_LINE_CALLOUT_2_NO_BORDER = 'callout2';
    public const TYPE_LINE_CALLOUT_3 = 'borderCallout3';
    public const TYPE_LINE_CALLOUT_3_ACCENT_BAR = 'accentCallout3';
    public const TYPE_LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR = 'accentBorderCallout3';
    public const TYPE_LINE_CALLOUT_3_NO_BORDER = 'callout3';
    public const TYPE_LINE_CALLOUT_4 = 'borderCallout4';
    public const TYPE_LINE_CALLOUT_4_ACCENT_BAR = 'accentCallout4';
    public const TYPE_LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR = 'accentBorderCallout4';
    public const TYPE_LINE_CALLOUT_4_NO_BORDER = 'callout4';
    public const TYPE_LINE_INVERSE = 'lineInv';
    public const TYPE_MATH_DIVIDE = 'mathDivide';
    public const TYPE_MATH_EQUAL = 'mathEqual';
    public const TYPE_MATH_MINUS = 'mathMinus';
    public const TYPE_MATH_MULTIPLY = 'mathMultiply';
    public const TYPE_MATH_NOT_EQUAL = 'mathNotEqual';
    public const TYPE_MATH_PLUS = 'mathPlus';
    //public const TYPE_MIXED = '';
    public const TYPE_MOON = 'moon';
    public const TYPE_NON_ISOSCELES_TRAPEZOID = 'nonIsoscelesTrapezoid';
    public const TYPE_NO_SYMBOL = 'noSmoking';
    public const TYPE_NOTCHED_RIGHT_ARROW = 'notchedRightArrow';
    //public const TYPE_NOTPRIMITIVE = '';
    public const TYPE_OCTAGON = 'octagon';
    public const TYPE_OVAL = 'ellipse';
    public const TYPE_OVAL_CALLOUT = 'wedgeEllipseCallout';
    public const TYPE_PARALLELOGRAM = 'parallelogram';
    public const TYPE_PENTAGON = 'homePlate';
    public const TYPE_PIE = 'pie';
    public const TYPE_PIE_WEDGE = 'pieWedge';
    public const TYPE_PLAQUE = 'plaque';
    public const TYPE_PLAQUE_TABS = 'plaqueTabs';
    public const TYPE_QUAD_ARROW = 'quadArrow';
    public const TYPE_QUAD_ARROW_CALLOUT = 'quadArrowCallout';
    public const TYPE_RECTANGLE = 'rect';
    public const TYPE_RECTANGULAR_CALLOUT = 'wedgeRectCallout';
    public const TYPE_REGULARP_ENTAGON = 'pentagon';
    public const TYPE_RIGHT_ARROW = 'rightArrow';
    public const TYPE_RIGHT_ARROW_CALLOUT = 'rightArrowCallout';
    public const TYPE_RIGHT_BRACE = 'rightBrace';
    public const TYPE_RIGHT_BRACKET = 'rightBracket';
    public const TYPE_RIGHT_TRIANGLE = 'rtTriangle';
    public const TYPE_ROUND_1_RECTANGLE = 'round1Rect';
    public const TYPE_ROUND_2_DIAG_RECTANGLE = 'round2DiagRect';
    public const TYPE_ROUND_2_SAME_RECTANGLE = 'round2SameRect';
    public const TYPE_ROUNDED_RECTANGLE = 'roundRect';
    public const TYPE_ROUNDED_RECTANGULAR_CALLOUT = 'wedgeRoundRectCallout';
    public const TYPE_SMILEY_FACE = 'smileyFace';
    public const TYPE_SNIP_1_RECTANGLE = 'snip1Rect';
    public const TYPE_SNIP_2_DIAG_RECTANGLE = 'snip2DiagRect';
    public const TYPE_SNIP_2_SAME_RECTANGLE = 'snip2SameRect';
    public const TYPE_SNIP_ROUND_RECTANGLE = 'snipRoundRect';
    public const TYPE_SQUARE_TABS = 'squareTabs';
    public const TYPE_STRIPED_RIGHT_ARROW = 'stripedRightArrow';
    public const TYPE_SUN = 'sun';
    public const TYPE_SWOOSH_ARROW = 'swooshArrow';
    public const TYPE_TEAR = 'teardrop';
    public const TYPE_TRAPEZOID = 'trapezoid';
    public const TYPE_UP_ARROW = 'upArrow';
    public const TYPE_UP_ARROW_CALLOUT = 'upArrowCallout';
    public const TYPE_UP_DOWN_ARROW = 'upDownArrow';
    public const TYPE_UP_DOWN_ARROW_CALLOUT = 'upDownArrowCallout';
    public const TYPE_UP_RIBBON = 'ribbon2';
    public const TYPE_U_TURN_ARROW = 'uturnArrow';
    public const TYPE_VERTICAL_SCROLL = 'verticalScroll';
    public const TYPE_WAVE = 'wave';

    /**
     * @var string
     */
    protected $text = '';

    /**
     * @var string
     */
    protected $type = self::TYPE_HEART;

    /**
     * @var Outline
     */
    protected $outline;

    public function __construct()
    {
        parent::__construct();

        $this->outline = new Outline();
    }

    public function getText(): string
    {
        return $this->text;
    }

    public function setText(string $text): self
    {
        $this->text = $text;

        return $this;
    }

    public function getType(): string
    {
        return $this->type;
    }

    public function setType(string $type): self
    {
        $this->type = $type;

        return $this;
    }

    public function getOutline(): Outline
    {
        return $this->outline;
    }

    public function setOutline(Outline $outline): self
    {
        $this->outline = $outline;

        return $this;
    }
}
