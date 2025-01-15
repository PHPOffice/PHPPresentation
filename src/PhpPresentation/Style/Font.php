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
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\Exception\NotAllowedValueException;

/**
 * \PhpOffice\PhpPresentation\Style\Font.
 */
class Font implements ComparableInterface
{
    // Capitalization type
    public const CAPITALIZATION_NONE = 'none';
    public const CAPITALIZATION_SMALL = 'small';
    public const CAPITALIZATION_ALL = 'all';

    // Charset type
    public const CHARSET_DEFAULT = 0x01;

    // Format type
    public const FORMAT_LATIN = 'latin';
    public const FORMAT_EAST_ASIAN = 'ea';
    public const FORMAT_COMPLEX_SCRIPT = 'cs';

    // Strike type
    public const STRIKE_NONE = 'noStrike';
    public const STRIKE_SINGLE = 'sngStrike';
    public const STRIKE_DOUBLE = 'dblStrike';

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

    // Script sub and super values
    public const BASELINE_SUPERSCRIPT = 300000;
    public const BASELINE_SUBSCRIPT = -250000;

    /**
     * Name.
     *
     * @var string
     */
    private $name = 'Calibri';

    /**
     * Panose.
     *
     * @var string
     */
    private $panose = '';

    /**
     * Pitch Family.
     *
     * @var int
     */
    private $pitchFamily = 0;

    /**
     * Charset.
     *
     * @var int
     */
    private $charset = self::CHARSET_DEFAULT;

    /**
     * Font Size.
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
     * Baseline.
     *
     * @var int
     */
    private $baseline = 0;

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
     * @var string
     */
    private $strikethrough = self::STRIKE_NONE;

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
     * Get panose.
     */
    public function getPanose(): string
    {
        return $this->panose;
    }

    /**
     * Set panose.
     */
    public function setPanose(string $pValue): self
    {
        if (mb_strlen($pValue) === 20) {
            $pValue = preg_replace('/.(.)/', '$1', $pValue);
        }

        if (mb_strlen($pValue) !== 10) {
            throw new InvalidParameterException('pValue', $pValue, 'The length is not correct');
        }

        $allowedChars = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'];
        foreach (mb_str_split($pValue) as $char) {
            if (!in_array($char, $allowedChars)) {
                throw new InvalidParameterException(
                    'pValue',
                    $pValue,
                    sprintf('The character "%s" is not allowed', $char)
                );
            }
        }

        $this->panose = $pValue;

        return $this;
    }

    /**
     * Get pitchFamily.
     */
    public function getPitchFamily(): int
    {
        return $this->pitchFamily;
    }

    /**
     * Set pitchFamily.
     */
    public function setPitchFamily(int $pValue): self
    {
        $this->pitchFamily = $pValue;

        return $this;
    }

    /**
     * Get charset.
     */
    public function getCharset(): int
    {
        return $this->charset;
    }

    /**
     * Set charset.
     */
    public function setCharset(int $pValue): self
    {
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
     * Set Baseline.
     */
    public function setBaseline(int $pValue): self
    {
        $this->baseline = $pValue;

        return $this;
    }

    /**
     * Get Baseline.
     */
    public function getBaseline(): int
    {
        return $this->baseline;
    }

    /**
     * Get SuperScript.
     *
     * @deprecated getBaseline() === self::BASELINE_SUPERSCRIPT
     */
    public function isSuperScript(): bool
    {
        return $this->getBaseline() === self::BASELINE_SUPERSCRIPT;
    }

    /**
     * Set SuperScript.
     *
     * @deprecated setBaseline(self::BASELINE_SUPERSCRIPT)
     */
    public function setSuperScript(bool $pValue = false): self
    {
        return $this->setBaseline($pValue ? self::BASELINE_SUPERSCRIPT : ($this->getBaseline() == self::BASELINE_SUBSCRIPT ? $this->getBaseline() : 0));
    }

    /**
     * Get SubScript.
     *
     * @deprecated getBaseline() === self::BASELINE_SUBSCRIPT
     */
    public function isSubScript(): bool
    {
        return $this->getBaseline() === self::BASELINE_SUBSCRIPT;
    }

    /**
     * Set SubScript.
     *
     * @deprecated setBaseline(self::BASELINE_SUBSCRIPT)
     */
    public function setSubScript(bool $pValue = false): self
    {
        return $this->setBaseline($pValue ? self::BASELINE_SUBSCRIPT : ($this->getBaseline() == self::BASELINE_SUPERSCRIPT ? $this->getBaseline() : 0));
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
     *
     * @deprecated Use `getStrikethrough`
     */
    public function isStrikethrough(): bool
    {
        return $this->strikethrough !== self::STRIKE_NONE;
    }

    /**
     * Get Strikethrough.
     */
    public function getStrikethrough(): string
    {
        return $this->strikethrough;
    }

    /**
     * Set Strikethrough.
     *
     * @deprecated $pValue as boolean
     *
     * @param bool|string $pValue
     *
     * @return self
     */
    public function setStrikethrough($pValue = false)
    {
        if (is_bool($pValue)) {
            $pValue = $pValue ? self::STRIKE_SINGLE : self::STRIKE_NONE;
        }
        if (in_array($pValue, [
            self::STRIKE_NONE,
            self::STRIKE_SINGLE,
            self::STRIKE_DOUBLE,
        ])) {
            $this->strikethrough = $pValue;
        }

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
            . $this->baseline
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
