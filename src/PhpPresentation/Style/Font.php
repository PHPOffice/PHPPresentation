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

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\Style\Font
 */
class Font implements ComparableInterface
{
    /* Underline types */
    public const UNDERLINE_NONE = 'none';
    public const UNDERLINE_DASH = 'dash';
    public const UNDERLINE_DASHHEAVY = 'dashHeavy';
    public const UNDERLINE_DASHLONG = 'dashLong';
    public const UNDERLINE_DASHLONGHEAVY = 'dashLongHeavy';
    public const UNDERLINE_DOUBLE = 'dbl';
    public const UNDERLINE_DOTHASH = 'dotDash';
    public const UNDERLINE_DOTHASHHEAVY = 'dotDashHeavy';
    public const UNDERLINE_DOTDOTDASH = 'dotDotDash';
    public const UNDERLINE_DOTDOTDASHHEAVY = 'dotDotDashHeavy';
    public const UNDERLINE_DOTTED = 'dotted';
    public const UNDERLINE_DOTTEDHEAVY = 'dottedHeavy';
    public const UNDERLINE_HEAVY = 'heavy';
    public const UNDERLINE_SINGLE = 'sng';
    public const UNDERLINE_WAVY = 'wavy';
    public const UNDERLINE_WAVYDOUBLE = 'wavyDbl';
    public const UNDERLINE_WAVYHEAVY = 'wavyHeavy';
    public const UNDERLINE_WORDS = 'words';

    /**
     * Name
     *
     * @var string
     */
    private $name;

    /**
     * Font Size
     *
     * @var int
     */
    private $size;

    /**
     * Bold
     *
     * @var boolean
     */
    private $bold;

    /**
     * Italic
     *
     * @var boolean
     */
    private $italic;

    /**
     * Superscript
     *
     * @var boolean
     */
    private $superScript;

    /**
     * Subscript
     *
     * @var boolean
     */
    private $subScript;

    /**
     * Underline
     *
     * @var string
     */
    private $underline;

    /**
     * Strikethrough
     *
     * @var boolean
     */
    private $strikethrough;

    /**
     * Foreground color
     *
     * @var Color
     */
    private $color;

    /**
     * Character Spacing
     *
     * @var int
     */
    private $characterSpacing;

    /**
     * Hash index
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Font
     */
    public function __construct()
    {
        // Initialise values
        $this->name             = 'Calibri';
        $this->size             = 10;
        $this->characterSpacing = 0;
        $this->bold             = false;
        $this->italic           = false;
        $this->superScript      = false;
        $this->subScript        = false;
        $this->underline        = self::UNDERLINE_NONE;
        $this->strikethrough    = false;
        $this->color            = new Color(Color::COLOR_BLACK);
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Set Name
     *
     * @param string                   $pValue
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setName($pValue = 'Calibri')
    {
        if ($pValue == '') {
            $pValue = 'Calibri';
        }
        $this->name = $pValue;

        return $this;
    }

    /**
     * Get Character Spacing
     *
     * @return double
     */
    public function getCharacterSpacing()
    {
        return $this->characterSpacing;
    }

    /**
     * Set Character Spacing
     * Value in pt
     * @param float|int $pValue
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setCharacterSpacing($pValue = 0)
    {
        if ($pValue == '') {
            $pValue = 0;
        }
        $this->characterSpacing = $pValue * 100;

        return $this;
    }

    /**
     * Get Size
     *
     * @return int
     */
    public function getSize(): int
    {
        return $this->size;
    }

    /**
     * Set Size
     *
     * @param int $pValue
     * @return self
     */
    public function setSize(int $pValue = 10): self
    {
        $this->size = $pValue;

        return $this;
    }

    /**
     * Get Bold
     *
     * @return boolean
     */
    public function isBold()
    {
        return $this->bold;
    }

    /**
     * Set Bold
     *
     * @param bool $pValue
     * @return self
     */
    public function setBold(bool $pValue = false): self
    {
        $this->bold = $pValue;

        return $this;
    }

    /**
     * Get Italic
     *
     * @return boolean
     */
    public function isItalic()
    {
        return $this->italic;
    }

    /**
     * Set Italic
     *
     * @param bool $pValue
     * @return self
     */
    public function setItalic(bool $pValue = false): self
    {
        $this->italic = $pValue;

        return $this;
    }

    /**
     * Get SuperScript
     *
     * @return boolean
     */
    public function isSuperScript()
    {
        return $this->superScript;
    }

    /**
     * Set SuperScript
     *
     * @param bool $pValue
     * @return self
     */
    public function setSuperScript(bool $pValue = false): self
    {
        $this->superScript = $pValue;

        // Set SubScript at false only if SuperScript is true
        if ($pValue === true) {
            $this->subScript = false;
        }

        return $this;
    }

    /**
     * @return bool
     */
    public function isSubScript(): bool
    {
        return $this->subScript;
    }

    /**
     * @param bool $pValue
     * @return self
     */
    public function setSubScript(bool $pValue = false): self
    {
        $this->subScript = $pValue;

        // Set SuperScript at false only if SubScript is true
        if ($pValue === true) {
            $this->superScript = false;
        }

        return $this;
    }

    /**
     * Get Underline
     *
     * @return string
     */
    public function getUnderline()
    {
        return $this->underline;
    }

    /**
     * Set Underline
     *
     * @param string                   $pValue \PhpOffice\PhpPresentation\Style\Font underline type
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setUnderline($pValue = self::UNDERLINE_NONE)
    {
        if ($pValue == '') {
            $pValue = self::UNDERLINE_NONE;
        }
        $this->underline = $pValue;

        return $this;
    }

    /**
     * Get Strikethrough
     *
     * @return boolean
     */
    public function isStrikethrough()
    {
        return $this->strikethrough;
    }

    /**
     * Set Strikethrough
     *
     * @param bool $pValue
     * @return self
     */
    public function setStrikethrough(bool $pValue = false): self
    {
        $this->strikethrough = $pValue;

        return $this;
    }

    /**
     * Get Color
     *
     * @return Color
     */
    public function getColor(): Color
    {
        return $this->color;
    }

    /**
     * Set Color
     *
     * @param Color $pValue
     * @throws \Exception
     * @return self
     */
    public function setColor(Color $pValue): self
    {
        $this->color = $pValue;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5($this->name . $this->size . ($this->bold ? 't' : 'f') . ($this->italic ? 't' : 'f') . ($this->superScript ? 't' : 'f') . ($this->subScript ? 't' : 'f') . $this->underline . ($this->strikethrough ? 't' : 'f') . $this->color->getHashCode() . __CLASS__);
    }

    /**
     * Get hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return int|null Hash index
     */
    public function getHashIndex(): ?int
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param int $value Hash index
     * @return $this
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;
        return $this;
    }
}
