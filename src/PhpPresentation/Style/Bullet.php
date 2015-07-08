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
 * \PhpOffice\PhpPresentation\Style\Bullet
 */
class Bullet implements ComparableInterface
{
    /* Bullet types */
    const TYPE_NONE                         = 'none';
    const TYPE_BULLET                       = 'bullet';
    const TYPE_NUMERIC                      = 'numeric';

    /* Numeric bullet styles */
    const NUMERIC_DEFAULT                   = 'arabicPeriod';
    const NUMERIC_ALPHALCPARENBOTH          = 'alphaLcParenBoth';
    const NUMERIC_ALPHAUCPARENBOTH          = 'alphaUcParenBoth';
    const NUMERIC_ALPHALCPARENR             = 'alphaLcParenR';
    const NUMERIC_ALPHAUCPARENR             = 'alphaUcParenR';
    const NUMERIC_ALPHALCPERIOD             = 'alphaLcPeriod';
    const NUMERIC_ALPHAUCPERIOD             = 'alphaUcPeriod';
    const NUMERIC_ARABICPARENBOTH           = 'arabicParenBoth';
    const NUMERIC_ARABICPARENR              = 'arabicParenR';
    const NUMERIC_ARABICPERIOD              = 'arabicPeriod';
    const NUMERIC_ARABICPLAIN               = 'arabicPlain';
    const NUMERIC_ROMANLCPARENBOTH          = 'romanLcParenBoth';
    const NUMERIC_ROMANUCPARENBOTH          = 'romanUcParenBoth';
    const NUMERIC_ROMANLCPARENR             = 'romanLcParenR';
    const NUMERIC_ROMANUCPARENR             = 'romanUcParenR';
    const NUMERIC_ROMANLCPERIOD             = 'romanLcPeriod';
    const NUMERIC_ROMANUCPERIOD             = 'romanUcPeriod';
    const NUMERIC_CIRCLENUMDBPLAIN          = 'circleNumDbPlain';
    const NUMERIC_CIRCLENUMWDBLACKPLAIN     = 'circleNumWdBlackPlain';
    const NUMERIC_CIRCLENUMWDWHITEPLAIN     = 'circleNumWdWhitePlain';
    const NUMERIC_ARABICDBPERIOD            = 'arabicDbPeriod';
    const NUMERIC_ARABICDBPLAIN             = 'arabicDbPlain';
    const NUMERIC_EA1CHSPERIOD              = 'ea1ChsPeriod';
    const NUMERIC_EA1CHSPLAIN               = 'ea1ChsPlain';
    const NUMERIC_EA1CHTPERIOD              = 'ea1ChtPeriod';
    const NUMERIC_EA1CHTPLAIN               = 'ea1ChtPlain';
    const NUMERIC_EA1JPNCHSDBPERIOD         = 'ea1JpnChsDbPeriod';
    const NUMERIC_EA1JPNKORPLAIN            = 'ea1JpnKorPlain';
    const NUMERIC_EA1JPNKORPERIOD           = 'ea1JpnKorPeriod';
    const NUMERIC_ARABIC1MINUS              = 'arabic1Minus';
    const NUMERIC_ARABIC2MINUS              = 'arabic2Minus';
    const NUMERIC_HEBREW2MINUS              = 'hebrew2Minus';
    const NUMERIC_THAIALPHAPERIOD           = 'thaiAlphaPeriod';
    const NUMERIC_THAIALPHAPARENR           = 'thaiAlphaParenR';
    const NUMERIC_THAIALPHAPARENBOTH        = 'thaiAlphaParenBoth';
    const NUMERIC_THAINUMPERIOD             = 'thaiNumPeriod';
    const NUMERIC_THAINUMPARENR             = 'thaiNumParenR';
    const NUMERIC_THAINUMPARENBOTH          = 'thaiNumParenBoth';
    const NUMERIC_HINDIALPHAPERIOD          = 'hindiAlphaPeriod';
    const NUMERIC_HINDINUMPERIOD            = 'hindiNumPeriod';
    const NUMERIC_HINDINUMPARENR            = 'hindiNumParenR';
    const NUMERIC_HINDIALPHA1PERIOD         = 'hindiAlpha1Period';

    /**
     * Bullet type
     *
     * @var string
     */
    private $bulletType = self::TYPE_NONE;

    /**
     * Bullet font
     *
     * @var string
     */
    private $bulletFont;

    /**
     * Bullet char
     *
     * @var string
     */
    private $bulletChar = '-';

    /**
     * Bullet numeric style
     *
     * @var string
     */
    private $bulletNumericStyle = self::NUMERIC_DEFAULT;

    /**
     * Bullet numeric start at
     *
     * @var int
     */
    private $bulletNumericStartAt = 1;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Bullet
     */
    public function __construct()
    {
        // Initialise values
        $this->bulletType              = self::TYPE_NONE;
        $this->bulletFont              = 'Calibri';
        $this->bulletChar              = '-';
        $this->bulletNumericStyle      = self::NUMERIC_DEFAULT;
        $this->bulletNumericStartAt    = 1;
    }

    /**
     * Get bullet type
     *
     * @return string
     */
    public function getBulletType()
    {
        return $this->bulletType;
    }

    /**
     * Set bullet type
     *
     * @param  string                     $pValue
     * @return \PhpOffice\PhpPresentation\Style\Bullet
     */
    public function setBulletType($pValue = self::TYPE_NONE)
    {
        $this->bulletType = $pValue;

        return $this;
    }

    /**
     * Get bullet font
     *
     * @return string
     */
    public function getBulletFont()
    {
        return $this->bulletFont;
    }

    /**
     * Set bullet font
     *
     * @param  string                     $pValue
     * @return \PhpOffice\PhpPresentation\Style\Bullet
     */
    public function setBulletFont($pValue = 'Calibri')
    {
        if ($pValue == '') {
            $pValue = 'Calibri';
        }
        $this->bulletFont = $pValue;

        return $this;
    }

    /**
     * Get bullet char
     *
     * @return string
     */
    public function getBulletChar()
    {
        return $this->bulletChar;
    }

    /**
     * Set bullet char
     *
     * @param  string                     $pValue
     * @return \PhpOffice\PhpPresentation\Style\Bullet
     */
    public function setBulletChar($pValue = '-')
    {
        $this->bulletChar = $pValue;

        return $this;
    }

    /**
     * Get bullet numeric style
     *
     * @return string
     */
    public function getBulletNumericStyle()
    {
        return $this->bulletNumericStyle;
    }

    /**
     * Set bullet numeric style
     *
     * @param  string                     $pValue
     * @return \PhpOffice\PhpPresentation\Style\Bullet
     */
    public function setBulletNumericStyle($pValue = self::NUMERIC_DEFAULT)
    {
        $this->bulletNumericStyle = $pValue;

        return $this;
    }

    /**
     * Get bullet numeric start at
     *
     * @return string
     */
    public function getBulletNumericStartAt()
    {
        return $this->bulletNumericStartAt;
    }

    /**
     * Set bullet numeric start at
     *
     * @param int|string $pValue
     * @return \PhpOffice\PhpPresentation\Style\Bullet
     */
    public function setBulletNumericStartAt($pValue = 1)
    {
        $this->bulletNumericStartAt = $pValue;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5(
            $this->bulletType
            . $this->bulletFont
            . $this->bulletChar
            . $this->bulletNumericStyle
            . $this->bulletNumericStartAt
            . __CLASS__
        );
    }

    /**
     * Get hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return string Hash index
     */
    public function getHashIndex()
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param string $value Hash index
     */
    public function setHashIndex($value)
    {
        $this->hashIndex = $value;
    }
}
