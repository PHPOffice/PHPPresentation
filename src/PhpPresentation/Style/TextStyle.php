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
            $oColorLT1 = new SchemeColor();
            $oColorLT1->setValue('lt1');
            $oColorTX1 = new SchemeColor();
            $oColorTX1->setValue('tx1');

            $oRTParagraphBody = new RichTextParagraph();
            $oRTParagraphBody->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setIndent(-324900 / 9525)
                ->setMarginLeft(342900 / 9525);
            $oRTParagraphBody->getFont()->setSize(32)->setColor($oColorTX1);
            $this->bodyStyle[1] = $oRTParagraphBody;

            $oRTParagraphOther = new RichTextParagraph();
            $oRTParagraphOther->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $oRTParagraphOther->getFont()->setSize(10)->setColor($oColorTX1);
            $this->otherStyle[0] = $oRTParagraphOther;

            $oRTParagraphTitle = new RichTextParagraph();
            $oRTParagraphTitle->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $oRTParagraphTitle->getFont()->setSize(44)->setColor($oColorLT1);
            $this->titleStyle[1] = $oRTParagraphTitle;
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
     * @return TextStyle
     */
    public function setBodyStyleAtLvl(RichTextParagraph $style, $lvl)
    {
        if ($this->checkLvl($lvl)) {
            $this->bodyStyle[$lvl] = $style;
        }
        return $this;
    }

    /**
     * @param RichTextParagraph $style
     * @param $lvl
     * @return TextStyle
     */
    public function setTitleStyleAtLvl(RichTextParagraph $style, $lvl)
    {
        if ($this->checkLvl($lvl)) {
            $this->titleStyle[$lvl] = $style;
        }
        return $this;
    }

    /**
     * @param RichTextParagraph $style
     * @param $lvl
     * @return TextStyle
     */
    public function setOtherStyleAtLvl(RichTextParagraph $style, $lvl)
    {
        if ($this->checkLvl($lvl)) {
            $this->otherStyle[$lvl] = $style;
        }
        return $this;
    }

    /**
     * @param $lvl
     * @return mixed
     */
    public function getBodyStyleAtLvl($lvl)
    {
        if ($this->checkLvl($lvl) && !empty($this->bodyStyle[$lvl])) {
            return $this->bodyStyle[$lvl];
        }
        return null;
    }

    /**
     * @param $lvl
     * @return mixed
     */
    public function getTitleStyleAtLvl($lvl)
    {
        if ($this->checkLvl($lvl) && !empty($this->titleStyle[$lvl])) {
            return $this->titleStyle[$lvl];
        }
        return null;
    }

    /**
     * @param $lvl
     * @return mixed
     */
    public function getOtherStyleAtLvl($lvl)
    {
        if ($this->checkLvl($lvl) && !empty($this->otherStyle[$lvl])) {
            return $this->otherStyle[$lvl];
        }
        return null;
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
}
