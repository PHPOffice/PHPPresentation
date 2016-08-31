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
}
