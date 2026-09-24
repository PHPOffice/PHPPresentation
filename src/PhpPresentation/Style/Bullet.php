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

/**
 * \PhpOffice\PhpPresentation\Style\Bullet.
 */
class Bullet implements ComparableInterface
{
    // Bullet types
    public const TYPE_NONE = 'none';
    public const TYPE_BULLET = 'bullet';
    public const TYPE_NUMERIC = 'numeric';

    // Numeric bullet styles
    public const NUMERIC_DEFAULT = 'arabicPeriod';
    public const NUMERIC_ALPHALCPARENBOTH = 'alphaLcParenBoth';
    public const NUMERIC_ALPHAUCPARENBOTH = 'alphaUcParenBoth';
    public const NUMERIC_ALPHALCPARENR = 'alphaLcParenR';
    public const NUMERIC_ALPHAUCPARENR = 'alphaUcParenR';
    public const NUMERIC_ALPHALCPERIOD = 'alphaLcPeriod';
    public const NUMERIC_ALPHAUCPERIOD = 'alphaUcPeriod';
    public const NUMERIC_ARABICPARENBOTH = 'arabicParenBoth';
    public const NUMERIC_ARABICPARENR = 'arabicParenR';
    public const NUMERIC_ARABICPERIOD = 'arabicPeriod';
    public const NUMERIC_ARABICPLAIN = 'arabicPlain';
    public const NUMERIC_ROMANLCPARENBOTH = 'romanLcParenBoth';
    public const NUMERIC_ROMANUCPARENBOTH = 'romanUcParenBoth';
    public const NUMERIC_ROMANLCPARENR = 'romanLcParenR';
    public const NUMERIC_ROMANUCPARENR = 'romanUcParenR';
    public const NUMERIC_ROMANLCPERIOD = 'romanLcPeriod';
    public const NUMERIC_ROMANUCPERIOD = 'romanUcPeriod';
    public const NUMERIC_CIRCLENUMDBPLAIN = 'circleNumDbPlain';
    public const NUMERIC_CIRCLENUMWDBLACKPLAIN = 'circleNumWdBlackPlain';
    public const NUMERIC_CIRCLENUMWDWHITEPLAIN = 'circleNumWdWhitePlain';
    public const NUMERIC_ARABICDBPERIOD = 'arabicDbPeriod';
    public const NUMERIC_ARABICDBPLAIN = 'arabicDbPlain';
    public const NUMERIC_EA1CHSPERIOD = 'ea1ChsPeriod';
    public const NUMERIC_EA1CHSPLAIN = 'ea1ChsPlain';
    public const NUMERIC_EA1CHTPERIOD = 'ea1ChtPeriod';
    public const NUMERIC_EA1CHTPLAIN = 'ea1ChtPlain';
    public const NUMERIC_EA1JPNCHSDBPERIOD = 'ea1JpnChsDbPeriod';
    public const NUMERIC_EA1JPNKORPLAIN = 'ea1JpnKorPlain';
    public const NUMERIC_EA1JPNKORPERIOD = 'ea1JpnKorPeriod';
    public const NUMERIC_ARABIC1MINUS = 'arabic1Minus';
    public const NUMERIC_ARABIC2MINUS = 'arabic2Minus';
    public const NUMERIC_HEBREW2MINUS = 'hebrew2Minus';
    public const NUMERIC_THAIALPHAPERIOD = 'thaiAlphaPeriod';
    public const NUMERIC_THAIALPHAPARENR = 'thaiAlphaParenR';
    public const NUMERIC_THAIALPHAPARENBOTH = 'thaiAlphaParenBoth';
    public const NUMERIC_THAINUMPERIOD = 'thaiNumPeriod';
    public const NUMERIC_THAINUMPARENR = 'thaiNumParenR';
    public const NUMERIC_THAINUMPARENBOTH = 'thaiNumParenBoth';
    public const NUMERIC_HINDIALPHAPERIOD = 'hindiAlphaPeriod';
    public const NUMERIC_HINDINUMPERIOD = 'hindiNumPeriod';
    public const NUMERIC_HINDINUMPARENR = 'hindiNumParenR';
    public const NUMERIC_HINDIALPHA1PERIOD = 'hindiAlpha1Period';

    /**
     * Bullet type.
     *
     * @var string
     */
    private $bulletType = self::TYPE_NONE;

    /**
     * Bullet font.
     *
     * @var string
     */
    private $bulletFont;

    /**
     * Bullet char.
     *
     * @var string
     */
    private $bulletChar = '-';

    /**
     * Bullet char.
     *
     * @var Color
     */
    private $bulletColor;

    /**
     * Bullet numeric style.
     *
     * @var string
     */
    private $bulletNumericStyle = self::NUMERIC_DEFAULT;

    /**
     * Bullet numeric start at, or null when none is given.
     *
     * @var null|int
     */
    private $bulletNumericStartAt;

    /**
     * Whether a numbering given no start continues the one before it past a paragraph that ends it.
     *
     * @var bool
     */
    private $bulletNumericContinue = false;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    public function __construct()
    {
        $this->bulletType = self::TYPE_NONE;
        $this->bulletFont = 'Calibri';
        $this->bulletChar = '-';
        $this->bulletColor = new Color();
        $this->bulletNumericStyle = self::NUMERIC_DEFAULT;
        $this->bulletNumericStartAt = null;
    }

    /**
     * Get bullet type.
     *
     * @return string
     */
    public function getBulletType()
    {
        return $this->bulletType;
    }

    /**
     * Set bullet type.
     *
     * @param string $pValue
     *
     * @return Bullet
     */
    public function setBulletType($pValue = self::TYPE_NONE)
    {
        $this->bulletType = $pValue;

        return $this;
    }

    /**
     * Get bullet font.
     *
     * @return string
     */
    public function getBulletFont()
    {
        return $this->bulletFont;
    }

    /**
     * Set bullet font.
     *
     * @param string $pValue
     *
     * @return Bullet
     */
    public function setBulletFont($pValue = 'Calibri')
    {
        if ('' == $pValue) {
            $pValue = 'Calibri';
        }
        $this->bulletFont = $pValue;

        return $this;
    }

    /**
     * Get bullet char.
     *
     * @return string
     */
    public function getBulletChar()
    {
        return $this->bulletChar;
    }

    /**
     * Set bullet char.
     *
     * @param string $pValue
     *
     * @return Bullet
     */
    public function setBulletChar($pValue = '-')
    {
        $this->bulletChar = $pValue;

        return $this;
    }

    /**
     * Get bullet numeric style.
     *
     * @return string
     */
    public function getBulletNumericStyle()
    {
        return $this->bulletNumericStyle;
    }

    /**
     * Set bullet numeric style.
     *
     * @param string $pValue
     *
     * @return Bullet
     */
    public function setBulletNumericStyle($pValue = self::NUMERIC_DEFAULT)
    {
        $this->bulletNumericStyle = $pValue;

        return $this;
    }

    /**
     * Get bullet numeric start at, or null when none is given.
     *
     * @return null|int
     */
    public function getBulletNumericStartAt()
    {
        return $this->bulletNumericStartAt;
    }

    /**
     * Set bullet numeric start at.
     *
     * The start is a number whatever the scheme: 3 starts an `alphaUcPeriod` list at "C." and a
     * `romanUcPeriod` one at "III.". It is kept within 1 to 32767, the range OOXML allows
     * (`ST_TextBulletStartAtNum`), the way LibreOffice reads one that falls outside it.
     *
     * A paragraph given no start (null) continues the numbering of the paragraph before it at
     * its level and with its scheme, or starts at 1 when there is none; one given a start
     * begins a new numbering there. A start equal to the one the numbering before it began at
     * continues it all the same: OOXML has no other way to write it.
     *
     * @param null|int|string $pValue
     *
     * @return Bullet
     */
    public function setBulletNumericStartAt($pValue = 1)
    {
        $this->bulletNumericStartAt = null === $pValue ? null : max(1, min(32767, (int) $pValue));

        return $this;
    }

    /**
     * Whether the paragraph, when it is given no start and begins a numbering, continues the last
     * numbering of its level and scheme instead: 1, 2, a paragraph with no marker, then 3.
     */
    public function isBulletNumericContinue(): bool
    {
        return $this->bulletNumericContinue;
    }

    public function setBulletNumericContinue(bool $value = true): self
    {
        $this->bulletNumericContinue = $value;

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
            $this->bulletType
            . $this->bulletFont
            . $this->bulletChar
            . $this->bulletNumericStyle
            . $this->bulletNumericStartAt
            . ($this->bulletNumericContinue ? 'continue' : '')
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

    /**
     * @return Color
     */
    public function getBulletColor()
    {
        return $this->bulletColor;
    }

    /**
     * @return Bullet
     */
    public function setBulletColor(Color $bulletColor)
    {
        $this->bulletColor = $bulletColor;

        return $this;
    }
}
