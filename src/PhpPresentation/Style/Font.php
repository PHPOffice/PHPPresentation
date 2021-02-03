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

use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Exception\NotAllowedValueException;

/**
 * \PhpOffice\PhpPresentation\Style\Font.
 */
class Font implements ComparableInterface
{
    // Underline types
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
  
    /* Strike types */
    public const STRIKE_NONE = 'noStrike';
    public const STRIKE_SINGLE = 'sngStrike';
    public const STRIKE_DOUBLE = 'dblStrike';

    public const FORMAT_LATIN = 'latin';
    public const FORMAT_EAST_ASIAN = 'ea';
    public const FORMAT_COMPLEX_SCRIPT = 'cs';

    public const CAPITALIZATION_NONE = 'none';
    public const CAPITALIZATION_SMALL = 'small';
    public const CAPITALIZATION_ALL = 'all';
    
    /* Script sub and super values */
    const SCRIPT_SUPER = 30000;
    const SCRIPT_SUB = -25000;

    /**
     * Name.
     *
     * @var string
     */
    private $name = 'Calibri';

    /**
     * panose
     *
     * @var string
     */
    private $panose;
    /**
     * pitchFamily
     *
     * @var string
     */
    private $pitchFamily;
    /**
     * charset
     *
     * @var string
     */
    private $charset;
    
    /**
     * Font Size
     *
     * @var int
     */
    private $size = 10;

    /**
     * Bold.
     *
     * @var bool
     */
    private $bold = false;

    /**
     * Italic.
     *
     * @var bool
     */
    private $italic = false;

    /**
     * Superscript.
     *
     * @var bool
     */
    private $superScript = false;

    /**
     * Subscript.
     *
     * @var bool
     */
    private $subScript = false;

    /**
     * Capitalization.
     *
     * @var string
     */
    private $capitalization = self::CAPITALIZATION_NONE;

    /**
     * Underline.
     *
     * @var string
     */
    private $underline = self::UNDERLINE_NONE;

    /**
     * Strikethrough.
     *
     * @var bool
     */
    private $strikethrough = false;

    /**
     * Foreground color.
     *
     * @var Color
     */
    private $color;

    /**
     * Character Spacing.
     *
     * @var float
     */
    private $characterSpacing = 0;

    /**
     * Format.
     *
     * @var string
     */
    private $format = self::FORMAT_LATIN;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    public function __construct()
    {
        $this->color = new Color(Color::COLOR_BLACK);
        $this->superScript      = 0;
        $this->subScript        = 0;
        $this->strikethrough    = self::STRIKE_NONE;
    }

    /**
     * Get Name.
     */
    public function getName(): string
    {
        return $this->name;
    }

    /**
     * Set Name.
     */
    public function setName(string $pValue = 'Calibri'): self
    {
        if ('' == $pValue) {
            $pValue = 'Calibri';
        }
        $this->name = $pValue;
        return $this;
    }
    
    /**
     * Get panose
     *
     * @return string
     */
    public function getPanose()
    {
        return $this->panose;
    }

    /**
     * Set panose
     *
     * @param  string                   $pValue
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setPanose($pValue)
    {
        if ($pValue == '') {
            $pValue = '';
        }
        $this->panose = $pValue;

        return $this;
    }
    /**
     * Get pitchFamily
     *
     * @return string
     */
    public function getPitchFamily()
    {
        return $this->pitchFamily;
    }

    /**
     * Set pitchFamily
     *
     * @param  string                   $pValue
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setPitchFamily($pValue)
    {
        if ($pValue == '') {
            $pValue = '';
        }
        $this->pitchFamily = $pValue;

        return $this;
    }
    /**
     * Get charset
     *
     * @return string
     */
    public function getCharset()
    {
        return $this->charset;
    }

    /**
     * Set charset
     *
     * @param  string                   $pValue
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setCharset($pValue)
    {
        if ($pValue == '') {
            $pValue = '';
        }
        $this->charset = $pValue;

        return $this;
    }

    /**
     * Get Character Spacing.
     */
    public function getCharacterSpacing(): float
    {
        return $this->characterSpacing;
    }

    /**
     * Set Character Spacing
     * Value in pt.
     */
    public function setCharacterSpacing(float $pValue = 0): self
    {
        $this->characterSpacing = $pValue * 100;

        return $this;
    }

    /**
     * Get Size.
     */
    public function getSize(): int
    {
        return $this->size;
    }

    /**
     * Set Size.
     */
    public function setSize(int $pValue = 10): self
    {
        $this->size = $pValue;

        return $this;
    }

    /**
     * Get Bold.
     */
    public function isBold(): bool
    {
        return $this->bold;
    }

    /**
     * Set Bold.
     */
    public function setBold(bool $pValue = false): self
    {
        $this->bold = $pValue;

        return $this;
    }

    /**
     * Get Italic.
     */
    public function isItalic(): bool
    {
        return $this->italic;
    }

    /**
     * Set Italic.
     */
    public function setItalic(bool $pValue = false): self
    {
        $this->italic = $pValue;

        return $this;
    }

    /**
     * Get SuperScript.
     */
    public function isSuperScript(): int
    {
        return $this->superScript;
    }

    /**
     * Set SuperScript
     *
     * @param  integer               $pValue
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setSuperScript($pValue = 0)
    {
        if ($pValue == '') {
            $pValue = 0;
        }

        $this->superScript = $pValue;

        // Set SubScript at false only if SuperScript is true
        if ($pValue != 0) {
            $this->subScript = 0;
        }

        return $this;
    }

    /**
     * Get SubScript
     *
     * @return integer
     */
    public function isSubScript()
    {
        return $this->subScript;
    }

    /**
     * Set SubScript
     *
     * @param  integer                 $pValue
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setSubScript($pValue = 0)
    {
        if ($pValue == '') {
            $pValue = 0;
        }

        $this->subScript = $pValue;

        // Set SuperScript at false only if SubScript is true
        if ($pValue != 0) {
            $this->superScript = 0;
        }

        return $this;
    }

    /**
     * Get Capitalization.
     */
    public function getCapitalization(): string
    {
        return $this->capitalization;
    }

    /**
     * Set Capitalization.
     */
    public function setCapitalization(string $pValue = self::CAPITALIZATION_NONE): self
    {
        if (!in_array(
            $pValue,
            [self::CAPITALIZATION_NONE, self::CAPITALIZATION_ALL, self::CAPITALIZATION_SMALL]
        )) {
            throw new NotAllowedValueException($pValue, [self::CAPITALIZATION_NONE, self::CAPITALIZATION_ALL, self::CAPITALIZATION_SMALL]);
        }

        $this->capitalization = $pValue;

        return $this;
    }

    /**
     * Get Underline.
     */
    public function getUnderline(): string
    {
        return $this->underline;
    }

    /**
     * Set Underline.
     *
     * @param string $pValue Underline type
     */
    public function setUnderline(string $pValue = self::UNDERLINE_NONE): self
    {
        if ('' == $pValue) {
            $pValue = self::UNDERLINE_NONE;
        }
        $this->underline = $pValue;

        return $this;
    }

    /**
     * Get Strikethrough.
     */
    public function isStrikethrough(): bool
    {
        return $this->strikethrough;
    }

    /**
     * Set Strikethrough.
     */
    public function setStrikethrough($pValue = self::STRIKE_NONE)
    {
        if ($pValue == '') {
            $pValue = self::STRIKE_NONE;
        }
        $this->strikethrough = $pValue;

        return $this;
    }

    /**
     * Get Color.
     */
    public function getColor(): Color
    {
        return $this->color;
    }

    /**
     * Set Color.
     */
    public function setColor(Color $pValue): self
    {
        $this->color = $pValue;

        return $this;
    }

    /**
     * Get format.
     */
    public function getFormat(): string
    {
        return $this->format;
    }

    /**
     * Set format.
     */
    public function setFormat(string $value = self::FORMAT_LATIN): self
    {
        if (in_array($value, [
            self::FORMAT_COMPLEX_SCRIPT,
            self::FORMAT_EAST_ASIAN,
            self::FORMAT_LATIN,
        ])) {
            $this->format = $value;
        }

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5(
            $this->name
            . $this->size
            . ($this->bold ? 't' : 'f')
            . ($this->italic ? 't' : 'f')
            . ($this->superScript ? 't' : 'f')
            . ($this->subScript ? 't' : 'f')
            . $this->underline
            . ($this->strikethrough ? 't' : 'f')
            . $this->format
            . $this->color->getHashCode()
            . __CLASS__
        );
    }

    /**
     * Get hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return null|int Hash index
     */
    public function getHashIndex(): ?int
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param int $value Hash index
     *
     * @return $this
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;

        return $this;
    }
}
