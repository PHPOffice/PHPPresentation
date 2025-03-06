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

namespace PhpOffice\PhpPresentation\Style;

use PhpOffice\PhpPresentation\Shape\RichText\Paragraph as RichTextParagraph;

class TextStyle
{
    /**
     * @var array<int, RichTextParagraph>
     */
    protected $bodyStyle = [];

    /**
     * @var array<int, RichTextParagraph>
     */
    protected $titleStyle = [];

    /**
     * @var array<int, RichTextParagraph>
     */
    protected $otherStyle = [];

    /**
     * TextStyle constructor.
     */
    public function __construct(bool $default = true)
    {
        if ($default) {
            $oColorLT1 = new SchemeColor();
            $oColorLT1->setValue('lt1');
            $oColorTX1 = new SchemeColor();
            $oColorTX1->setValue('tx1');

            $oRTParagraphBody = new RichTextParagraph();
            $oRTParagraphBody->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setIndent(-342900 / 9525)
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

    private function checkLvl(?int $lvl): bool
    {
        if (null === $lvl || $lvl > 9) {
            return false;
        }

        return true;
    }

    public function setBodyStyleAtLvl(RichTextParagraph $style, ?int $lvl): self
    {
        if ($this->checkLvl($lvl)) {
            $this->bodyStyle[$lvl] = $style;
        }

        return $this;
    }

    public function setTitleStyleAtLvl(RichTextParagraph $style, ?int $lvl): self
    {
        if ($this->checkLvl($lvl)) {
            $this->titleStyle[$lvl] = $style;
        }

        return $this;
    }

    public function setOtherStyleAtLvl(RichTextParagraph $style, ?int $lvl): self
    {
        if ($this->checkLvl($lvl)) {
            $this->otherStyle[$lvl] = $style;
        }

        return $this;
    }

    public function getBodyStyleAtLvl(?int $lvl): ?RichTextParagraph
    {
        if ($this->checkLvl($lvl) && !empty($this->bodyStyle[$lvl])) {
            return $this->bodyStyle[$lvl];
        }

        return null;
    }

    public function getTitleStyleAtLvl(?int $lvl): ?RichTextParagraph
    {
        if ($this->checkLvl($lvl) && !empty($this->titleStyle[$lvl])) {
            return $this->titleStyle[$lvl];
        }

        return null;
    }

    public function getOtherStyleAtLvl(?int $lvl): ?RichTextParagraph
    {
        if ($this->checkLvl($lvl) && !empty($this->otherStyle[$lvl])) {
            return $this->otherStyle[$lvl];
        }

        return null;
    }

    /**
     * @return array<int, RichTextParagraph>
     */
    public function getBodyStyle(): array
    {
        return $this->bodyStyle;
    }

    /**
     * @return array<int, RichTextParagraph>
     */
    public function getTitleStyle(): array
    {
        return $this->titleStyle;
    }

    /**
     * @return array<int, RichTextParagraph>
     */
    public function getOtherStyle(): array
    {
        return $this->otherStyle;
    }
}
