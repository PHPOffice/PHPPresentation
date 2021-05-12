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

namespace PhpOffice\PhpPresentation\Shape;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * AutoShape shape
 *
 * @link : https://github.com/scanny/python-pptx/blob/eaa1e0fd3db28b03a353e116a5c7d2084dd87c26/pptx/enum/shapes.py
 */
class AutoShape extends AbstractShape implements ComparableInterface
{
    const TYPE_10_POINT_STAR = 'star10';
    const TYPE_12_POINT_STAR = 'star12';
    const TYPE_16_POINT_STAR = 'star16';
    const TYPE_24_POINT_STAR = 'star24';
    const TYPE_32_POINT_STAR = 'star32';
    const TYPE_4_POINT_STAR = 'star4';
    const TYPE_5_POINT_STAR = 'star5';
    const TYPE_6_POINT_STAR = 'star6';
    const TYPE_7_POINT_STAR = 'star7';
    const TYPE_8_POINT_STAR = 'star8';
    const TYPE_ACTION_BUTTON_BACK_OR_PREVIOUS = 'actionButtonBackPrevious';
    const TYPE_ACTION_BUTTON_BEGINNING = 'actionButtonBeginning';
    const TYPE_ACTION_BUTTON_CUSTOM = 'actionButtonBlank';
    const TYPE_ACTION_BUTTON_DOCUMENT = 'actionButtonDocument';
    const TYPE_ACTION_BUTTON_END = 'actionButtonEnd';
    const TYPE_ACTION_BUTTON_FORWARD_OR_NEXT = 'actionButtonForwardNext';
    const TYPE_ACTION_BUTTON_HELP = 'actionButtonHelp';
    const TYPE_ACTION_BUTTON_HOME = 'actionButtonHome';
    const TYPE_ACTION_BUTTON_INFORMATION = 'actionButtonInformation';
    const TYPE_ACTION_BUTTON_MOVIE = 'actionButtonMovie';
    const TYPE_ACTION_BUTTON_RETURN = 'actionButtonReturn';
    const TYPE_ACTION_BUTTON_SOUND = 'actionButtonSound';
    const TYPE_ARC = 'arc';
    const TYPE_BALLOON = 'wedgeRoundRectCallout';
    const TYPE_BENT_ARROW = 'bentArrow';
    const TYPE_BENT_UP_ARROW = 'bentUpArrow';
    const TYPE_BEVEL = 'bevel';
    const TYPE_BLOCK_ARC = 'blockArc';
    const TYPE_CAN = 'can';
    const TYPE_CHART_PLUS = 'chartPlus';
    const TYPE_CHART_STAR = 'chartStar';
    const TYPE_CHARTX = 'chartX';
    const TYPE_CHEVRON = 'chevron';
    const TYPE_CHORD = 'chord';
    const TYPE_CIRCULAR_ARROW = 'circularArrow';
    const TYPE_CLOUD = 'cloud';
    const TYPE_CLOUD_CALLOUT = 'cloudCallout';
    const TYPE_CORNER = 'corner';
    const TYPE_CORNER_TABS = 'cornerTabs';
    const TYPE_CROSS = 'plus';
    const TYPE_CUBE = 'cube';
    const TYPE_CURVED_DOWN_ARROW = 'curvedDownArrow';
    const TYPE_CURVED_DOWN_RIBBON = 'ellipseRibbon';
    const TYPE_CURVED_LEFT_ARROW = 'curvedLeftArrow';
    const TYPE_CURVED_RIGHT_ARROW = 'curvedRightArrow';
    const TYPE_CURVED_UP_ARROW = 'curvedUpArrow';
    const TYPE_CURVED_UP_RIBBON = 'ellipseRibbon2';
    const TYPE_DECAGON = 'decagon';
    const TYPE_DIAGONALSTRIPE = 'diagStripe';
    const TYPE_DIAMOND = 'diamond';
    const TYPE_DODECAGON = 'dodecagon';
    const TYPE_DONUT = 'donut';
    const TYPE_DOUBLE_BRACE = 'bracePair';
    const TYPE_DOUBLE_BRACKET = 'bracketPair';
    const TYPE_DOUBLE_WAVE = 'doubleWave';
    const TYPE_DOWN_ARROW = 'downArrow';
    const TYPE_DOWN_ARROWCALLOUT = 'downArrowCallout';
    const TYPE_DOWN_RIBBON = 'ribbon';
    const TYPE_EXPLOSIONEXPLOSION1 = 'irregularSeal1';
    const TYPE_EXPLOSIONEXPLOSION2 = 'irregularSeal2';
    const TYPE_FLOWCHART_ALTERNATEPROCESS = 'flowChartAlternateProcess';
    const TYPE_FLOWCHART_CARD = 'flowChartPunchedCard';
    const TYPE_FLOWCHART_COLLATE = 'flowChartCollate';
    const TYPE_FLOWCHART_CONNECTOR = 'flowChartConnector';
    const TYPE_FLOWCHART_DATA = 'flowChartInputOutput';
    const TYPE_FLOWCHART_DECISION = 'flowChartDecision';
    const TYPE_FLOWCHART_DELAY = 'flowChartDelay';
    const TYPE_FLOWCHART_DIRECT_ACCESS_STORAGE = 'flowChartMagneticDrum';
    const TYPE_FLOWCHART_DISPLAY = 'flowChartDisplay';
    const TYPE_FLOWCHART_DOCUMENT = 'flowChartDocument';
    const TYPE_FLOWCHART_EXTRACT = 'flowChartExtract';
    const TYPE_FLOWCHART_INTERNAL_STORAGE = 'flowChartInternalStorage';
    const TYPE_FLOWCHART_MAGNETIC_DISK = 'flowChartMagneticDisk';
    const TYPE_FLOWCHART_MANUAL_INPUT = 'flowChartManualInput';
    const TYPE_FLOWCHART_MANUAL_OPERATION = 'flowChartManualOperation';
    const TYPE_FLOWCHART_MERGE = 'flowChartMerge';
    const TYPE_FLOWCHART_MULTIDOCUMENT = 'flowChartMultidocument';
    const TYPE_FLOWCHART_OFFLINE_STORAGE = 'flowChartOfflineStorage';
    const TYPE_FLOWCHART_OFFPAGE_CONNECTOR = 'flowChartOffpageConnector';
    const TYPE_FLOWCHART_OR = 'flowChartOr';
    const TYPE_FLOWCHART_PREDEFINED_PROCESS = 'flowChartPredefinedProcess';
    const TYPE_FLOWCHART_PREPARATION = 'flowChartPreparation';
    const TYPE_FLOWCHART_PROCESS = 'flowChartProcess';
    const TYPE_FLOWCHART_PUNCHEDTAPE = 'flowChartPunchedTape';
    const TYPE_FLOWCHART_SEQUENTIAL_ACCESS_STORAGE = 'flowChartMagneticTape';
    const TYPE_FLOWCHART_SORT = 'flowChartSort';
    const TYPE_FLOWCHART_STORED_DATA = 'flowChartOnlineStorage';
    const TYPE_FLOWCHART_SUMMING_JUNCTION = 'flowChartSummingJunction';
    const TYPE_FLOWCHART_TERMINATOR = 'flowChartTerminator';
    const TYPE_FOLDED_CORNER = 'foldedCorner';
    const TYPE_FRAME = 'frame';
    const TYPE_FUNNEL = 'funnel';
    const TYPE_GEAR_6 = 'gear6';
    const TYPE_GEAR_9 = 'gear9';
    const TYPE_HALF_FRAME = 'halfFrame';
    const TYPE_HEART = 'heart';
    const TYPE_HEPTAGON = 'heptagon';
    const TYPE_HEXAGON = 'hexagon';
    const TYPE_HORIZONTAL_SCROLL = 'horizontalScroll';
    const TYPE_ISOSCELES_TRIANGLE = 'triangle';
    const TYPE_LEFT_ARROW = 'leftArrow';
    const TYPE_LEFT_ARROW_CALLOUT = 'leftArrowCallout';
    const TYPE_LEFT_BRACE = 'leftBrace';
    const TYPE_LEFT_BRACKET = 'leftBracket';
    const TYPE_LEFT_CIRCULAR_ARROW = 'leftCircularArrow';
    const TYPE_LEFT_RIGHT_ARROW = 'leftRightArrow';
    const TYPE_LEFT_RIGHT_ARROW_CALLOUT = 'leftRightArrowCallout';
    const TYPE_LEFT_RIGHT_CIRCULAR_ARROW = 'leftRightCircularArrow';
    const TYPE_LEFT_RIGHT_RIBBON = 'leftRightRibbon';
    const TYPE_LEFT_RIGHT_UP_ARROW = 'leftRightUpArrow';
    const TYPE_LEFT_UP_ARROW = 'leftUpArrow';
    const TYPE_LIGHTNING_BOLT = 'lightningBolt';
    const TYPE_LINE_CALLOUT_1 = 'borderCallout1';
    const TYPE_LINE_CALLOUT_1_ACCENT_BAR = 'accentCallout1';
    const TYPE_LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR = 'accentBorderCallout1';
    const TYPE_LINE_CALLOUT_1_NO_BORDER = 'callout1';
    const TYPE_LINE_CALLOUT_2 = 'borderCallout2';
    const TYPE_LINE_CALLOUT_2_ACCENT_BAR = 'accentCallout2';
    const TYPE_LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR = 'accentBorderCallout2';
    const TYPE_LINE_CALLOUT_2_NO_BORDER = 'callout2';
    const TYPE_LINE_CALLOUT_3 = 'borderCallout3';
    const TYPE_LINE_CALLOUT_3_ACCENT_BAR = 'accentCallout3';
    const TYPE_LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR = 'accentBorderCallout3';
    const TYPE_LINE_CALLOUT_3_NO_BORDER = 'callout3';
    const TYPE_LINE_CALLOUT_4 = 'borderCallout4';
    const TYPE_LINE_CALLOUT_4_ACCENT_BAR = 'accentCallout4';
    const TYPE_LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR = 'accentBorderCallout4';
    const TYPE_LINE_CALLOUT_4_NO_BORDER = 'callout4';
    const TYPE_LINE_INVERSE = 'lineInv';
    const TYPE_MATH_DIVIDE = 'mathDivide';
    const TYPE_MATH_EQUAL = 'mathEqual';
    const TYPE_MATH_MINUS = 'mathMinus';
    const TYPE_MATH_MULTIPLY = 'mathMultiply';
    const TYPE_MATH_NOT_EQUAL = 'mathNotEqual';
    const TYPE_MATH_PLUS = 'mathPlus';
    //const TYPE_MIXED = '';
    const TYPE_MOON = 'moon';
    const TYPE_NON_ISOSCELES_TRAPEZOID = 'nonIsoscelesTrapezoid';
    const TYPE_NO_SYMBOL = 'noSmoking';
    const TYPE_NOTCHED_RIGHT_ARROW = 'notchedRightArrow';
    //const TYPE_NOTPRIMITIVE = '';
    const TYPE_OCTAGON = 'octagon';
    const TYPE_OVAL = 'ellipse';
    const TYPE_OVAL_CALLOUT = 'wedgeEllipseCallout';
    const TYPE_PARALLELOGRAM = 'parallelogram';
    const TYPE_PENTAGON = 'homePlate';
    const TYPE_PIE = 'pie';
    const TYPE_PIE_WEDGE = 'pieWedge';
    const TYPE_PLAQUE = 'plaque';
    const TYPE_PLAQUE_TABS = 'plaqueTabs';
    const TYPE_QUAD_ARROW = 'quadArrow';
    const TYPE_QUAD_ARROW_CALLOUT = 'quadArrowCallout';
    const TYPE_RECTANGLE = 'rect';
    const TYPE_RECTANGULAR_CALLOUT = 'wedgeRectCallout';
    const TYPE_REGULARP_ENTAGON = 'pentagon';
    const TYPE_RIGHT_ARROW = 'rightArrow';
    const TYPE_RIGHT_ARROW_CALLOUT = 'rightArrowCallout';
    const TYPE_RIGHT_BRACE = 'rightBrace';
    const TYPE_RIGHT_BRACKET = 'rightBracket';
    const TYPE_RIGHT_TRIANGLE = 'rtTriangle';
    const TYPE_ROUND_1_RECTANGLE = 'round1Rect';
    const TYPE_ROUND_2_DIAG_RECTANGLE = 'round2DiagRect';
    const TYPE_ROUND_2_SAME_RECTANGLE = 'round2SameRect';
    const TYPE_ROUNDED_RECTANGLE = 'roundRect';
    const TYPE_ROUNDED_RECTANGULAR_CALLOUT = 'wedgeRoundRectCallout';
    const TYPE_SMILEY_FACE = 'smileyFace';
    const TYPE_SNIP_1_RECTANGLE = 'snip1Rect';
    const TYPE_SNIP_2_DIAG_RECTANGLE = 'snip2DiagRect';
    const TYPE_SNIP_2_SAME_RECTANGLE = 'snip2SameRect';
    const TYPE_SNIP_ROUND_RECTANGLE = 'snipRoundRect';
    const TYPE_SQUARE_TABS = 'squareTabs';
    const TYPE_STRIPED_RIGHT_ARROW = 'stripedRightArrow';
    const TYPE_SUN = 'sun';
    const TYPE_SWOOSH_ARROW = 'swooshArrow';
    const TYPE_TEAR = 'teardrop';
    const TYPE_TRAPEZOID = 'trapezoid';
    const TYPE_UP_ARROW = 'upArrow';
    const TYPE_UP_ARROW_CALLOUT = 'upArrowCallout';
    const TYPE_UP_DOWN_ARROW = 'upDownArrow';
    const TYPE_UP_DOWN_ARROW_CALLOUT = 'upDownArrowCallout';
    const TYPE_UP_RIBBON = 'ribbon2';
    const TYPE_U_TURN_ARROW = 'uturnArrow';
    const TYPE_VERTICAL_SCROLL = 'verticalScroll';
    const TYPE_WAVE = 'wave';

    /**
     * @var string
     */
    protected $text;

    /**
     * @var string
     */
    protected $type;

    public function __construct()
    {
        parent::__construct();
    }

    /**
     * @return string
     */
    public function getText()
    {
        return $this->text;
    }

    /**
     * @param string $text
     * @return AutoShape
     */
    public function setText($text)
    {
        $this->text = $text;
        return $this;
    }

    /**
     * @return string
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * @param string $type
     * @return AutoShape
     */
    public function setType($type)
    {
        $this->type = $type;
        return $this;
    }
}
