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

namespace PhpOffice\PhpPresentation\Style;

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph as RichTextParagraph;

/**
 * Class TextStyle
 */
class TextStyle
{
    /**
     * @var array
     */
    protected $bodyStyle = array();
    /**
     * @var array
     */
    protected $titleStyle = array();
    /**
     * @var array
     */
    protected $otherStyle = array();

    /**
     * TextStyle constructor.
     * @param bool $default
     */
    public function __construct($default = true)
    {
        if ($default) {
            $oRTParagraph = new RichTextParagraph();
            $oRTParagraph->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $oRTParagraph->getFont()->setSize(44)->setColor(new SchemeColor())->getColor()->setValue("lt1");
            $this->titleStyle[1] = $oRTParagraph;
            $oRTParagraph = new RichTextParagraph();
            $oRTParagraph->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setIndent(-324900 / 9525)
                ->setMarginLeft(342900 / 9525);
            $oRTParagraph->getFont()->setSize(32)->setColor(new SchemeColor())->getColor()->setValue("tx1");
            $this->bodyStyle[1] = $oRTParagraph;
            $oRTParagraph = new RichTextParagraph();
            $oRTParagraph->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $oRTParagraph->getFont()->setSize(10)->setColor(new SchemeColor())->getColor()->setValue("tx1");
            $this->otherStyle[0] = $oRTParagraph;
        }
    }

    /**
     * @param $lvl
     * @return bool
     */
    private function checkLvl($lvl)
    {
        if (!is_int($lvl)) {
            return false;
        }
        if ($lvl > 9) {
            return false;
        }
        return true;
    }

    /**
     * @param RichTextParagraph $style
     * @param $lvl
     */
    public function setBodyStyleAtLvl(RichTextParagraph $style, $lvl)
    {
        if (!$this->checkLvl($lvl)) {
            $this->bodyStyle[$lvl] = $style;
        }
    }

    /**
     * @param RichTextParagraph $style
     * @param $lvl
     */
    public function setTitleStyleAtLvl(RichTextParagraph $style, $lvl)
    {
        if (!$this->checkLvl($lvl)) {
            $this->titleStyle[$lvl] = $style;
        }
    }

    /**
     * @param RichTextParagraph $style
     * @param $lvl
     */
    public function setOtherStyleAtLvl(RichTextParagraph $style, $lvl)
    {
        if (!$this->checkLvl($lvl)) {
            $this->otherStyle[$lvl] = $style;
        }
    }

    /**
     * @param $lvl
     * @return mixed
     */
    public function getBodyStyleAtLvl($lvl)
    {
        if (!$this->checkLvl($lvl)) {
            return $this->bodyStyle[$lvl];
        }
    }

    /**
     * @param $lvl
     * @return mixed
     */
    public function getTitleStyleAtLvl($lvl)
    {
        if (!$this->checkLvl($lvl)) {
            return $this->bodyStyle[$lvl];
        }
    }

    /**
     * @param $lvl
     * @return mixed
     */
    public function getOtherStyleAtLvl($lvl)
    {
        if (!$this->checkLvl($lvl)) {
            return $this->bodyStyle[$lvl];
        }
    }

    /**
     * @return array
     */
    public function getBodyStyle()
    {
        return $this->bodyStyle;
    }

    /**
     * @return array
     */
    public function getTitleStyle()
    {
        return $this->titleStyle;
    }

    /**
     * @return array
     */
    public function getOtherStyle()
    {
        return $this->otherStyle;
    }
//
//
//
//
//
//
//
//
//
//
//
//<p:titleStyle>
//    <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
//        <a:spcBef>
//            <a:spcPct val="0" />
//        </a:spcBef>
//        <a:buNone />
//        <a:defRPr sz="4400" kern="1200">
//            <a:solidFill>
//                <a:schemeClr val="tx1" />
//            </a:solidFill>
//            <a:latin typeface="+mj-lt" />
//            <a:ea typeface="+mj-ea" />
//            <a:cs typeface="+mj-cs" />
//        </a:defRPr>
//    </a:lvl1pPr>
//</p:titleStyle>
//
//
//    // Create a shape (text)
//echo date('H:i:s') . ' Create a shape (rich text)'.EOL;
//$shape = $currentSlide->createRichTextShape();
//$shape->setHeight(600)
//->setWidth(930)
//->setOffsetX(10)
//->setOffsetY(130);
//$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
//->setMarginLeft(25)
//->setIndent(-25);
//$shape->getActiveParagraph()->getFont()->setSize(36)
//->setColor($colorBlack);
//$shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
//
//$shape->createTextRun('Generate slide decks');
//
//$shape->createParagraph()->getAlignment()->setLevel(1)
//->setMarginLeft(75)
//->setIndent(-25);
//$shape->createTextRun('Represent business data');
//$shape->createParagraph()->createTextRun('Show a family slide show');
//$shape->createParagraph()->createTextRun('...');
//
//$shape->createParagraph()->getAlignment()->setLevel(0)
//->setMarginLeft(25)
//->setIndent(-25);
//$shape->createTextRun('Export these to different formats');
//$shape->createParagraph()->getAlignment()->setLevel(1)
//->setMarginLeft(75)
//->setIndent(-25);
//$shape->createTextRun('PHPPresentation 2007');
//$shape->createParagraph()->createTextRun('Serialized');
//$shape->createParagraph()->createTextRun('... (more to come) ...');
//
    /**
     * Text Styles
     * p:bodyStyle
     * p:titleStyle
     * p:otherStyle
     *
     * a:defPPr
     * a:lvl1pPr
     * .
     * .
     * .
     * a:lvl9pPr
     * attributes
     * - algn (alignment)
     * - defTabSz (default tab size)
     * - fontAlgn (font alignment)
     * - hangingPunct (specifies the handling of hanging text)
     * - indent (indentation for the first line of text)
     * - latinLnBrk (specifies whether to break words)
     * - marL and marR (left and right margins)
     * - lang
     *elements*
     *
     * Bullets & Numbering (http://officeopenxml.com/drwSp-text-paraProps-numbering.php)
     * <a:buAutoNum/> (auto-numbering bullet)
     * <a:buBlip> (picture bullet)
     * <a:buChar/> (character bullet)
     * <a:buClr> (color of bullets)
     * <a:buClrTx> (color of bullets is same as text run)
     * <a:buFont/> (font for bullets)
     * <a:buFontTx> (font for bullets is same as text run)
     * <a:buNone> (no bullet)
     * <a:buSzPct> (size in percentage of bullet characters)
     * <a:buSzPts/> (size in points of bullet characters)
     * <a:buSzTx> (size of bullet characters to be size of text run)
     *
     * List Properties and Default Style. (http://officeopenxml.com/drwSp-text-lstPr.php)
     * <a:defRPr> (default text run properties)
     *
     * Text - Spacing, Indents and Margins. (http://officeopenxml.com/drwSp-text-paraProps-margins.php)
     * <a:lnSpc> (line spacing)
     * <a:spcAft> (spacing after the paragraph)
     * <a:spcBef> (spacing before the paragraph)
     *
     * Text - Alignment, Tabs, Other (http://officeopenxml.com/drwSp-text-paraProps-align.php)
     * <a:tabLst> (list of tab stops in a paragraph)
     */
}
