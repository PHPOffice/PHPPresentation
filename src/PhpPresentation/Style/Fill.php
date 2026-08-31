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

class Fill implements ComparableInterface
{
    // Fill types
    /**
     * No fill was asked for. Table cells and rows start here, so that a cell with nothing of its
     * own can be told apart from one deliberately left transparent, and painted with its row's
     * fill instead. Both formats make the same distinction: an empty `a:tcPr` against `a:noFill`,
     * an absent `table:style-name` against a cell style with `draw:fill="none"`.
     */
    public const FILL_UNSET = 'unset';
    public const FILL_NONE = 'none';
    public const FILL_SOLID = 'solid';
    public const FILL_GRADIENT_LINEAR = 'linear';
    public const FILL_GRADIENT_PATH = 'path';

    /**
     * The preset patterns of `ST_PresetPatternVal`, which is the vocabulary `a:pattFill` names one
     * by. All 54 of them; a pattern outside this list is not a pattern the format knows.
     */
    public const FILL_PATTERN_PCT5 = 'pct5';
    public const FILL_PATTERN_PCT10 = 'pct10';
    public const FILL_PATTERN_PCT20 = 'pct20';
    public const FILL_PATTERN_PCT25 = 'pct25';
    public const FILL_PATTERN_PCT30 = 'pct30';
    public const FILL_PATTERN_PCT40 = 'pct40';
    public const FILL_PATTERN_PCT50 = 'pct50';
    public const FILL_PATTERN_PCT60 = 'pct60';
    public const FILL_PATTERN_PCT70 = 'pct70';
    public const FILL_PATTERN_PCT75 = 'pct75';
    public const FILL_PATTERN_PCT80 = 'pct80';
    public const FILL_PATTERN_PCT90 = 'pct90';
    public const FILL_PATTERN_HORZ = 'horz';
    public const FILL_PATTERN_VERT = 'vert';
    public const FILL_PATTERN_LTHORZ = 'ltHorz';
    public const FILL_PATTERN_LTVERT = 'ltVert';
    public const FILL_PATTERN_DKHORZ = 'dkHorz';
    public const FILL_PATTERN_DKVERT = 'dkVert';
    public const FILL_PATTERN_NARHORZ = 'narHorz';
    public const FILL_PATTERN_NARVERT = 'narVert';
    public const FILL_PATTERN_DASHHORZ = 'dashHorz';
    public const FILL_PATTERN_DASHVERT = 'dashVert';
    public const FILL_PATTERN_CROSS = 'cross';
    public const FILL_PATTERN_DNDIAG = 'dnDiag';
    public const FILL_PATTERN_UPDIAG = 'upDiag';
    public const FILL_PATTERN_LTDNDIAG = 'ltDnDiag';
    public const FILL_PATTERN_LTUPDIAG = 'ltUpDiag';
    public const FILL_PATTERN_DKDNDIAG = 'dkDnDiag';
    public const FILL_PATTERN_DKUPDIAG = 'dkUpDiag';
    public const FILL_PATTERN_WDDNDIAG = 'wdDnDiag';
    public const FILL_PATTERN_WDUPDIAG = 'wdUpDiag';
    public const FILL_PATTERN_DASHDNDIAG = 'dashDnDiag';
    public const FILL_PATTERN_DASHUPDIAG = 'dashUpDiag';
    public const FILL_PATTERN_DIAGCROSS = 'diagCross';
    public const FILL_PATTERN_SMCHECK = 'smCheck';
    public const FILL_PATTERN_LGCHECK = 'lgCheck';
    public const FILL_PATTERN_SMGRID = 'smGrid';
    public const FILL_PATTERN_LGGRID = 'lgGrid';
    public const FILL_PATTERN_DOTGRID = 'dotGrid';
    public const FILL_PATTERN_SMCONFETTI = 'smConfetti';
    public const FILL_PATTERN_LGCONFETTI = 'lgConfetti';
    public const FILL_PATTERN_HORZBRICK = 'horzBrick';
    public const FILL_PATTERN_DIAGBRICK = 'diagBrick';
    public const FILL_PATTERN_SOLIDDMND = 'solidDmnd';
    public const FILL_PATTERN_OPENDMND = 'openDmnd';
    public const FILL_PATTERN_DOTDMND = 'dotDmnd';
    public const FILL_PATTERN_PLAID = 'plaid';
    public const FILL_PATTERN_SPHERE = 'sphere';
    public const FILL_PATTERN_WEAVE = 'weave';
    public const FILL_PATTERN_DIVOT = 'divot';
    public const FILL_PATTERN_SHINGLE = 'shingle';
    public const FILL_PATTERN_WAVE = 'wave';
    public const FILL_PATTERN_TRELLIS = 'trellis';
    public const FILL_PATTERN_ZIGZAG = 'zigZag';

    /**
     * The seventeen names this class carried before the preset patterns were written at all.
     *
     * Their values used to be the `ST_PatternType` of SpreadsheetML -- `darkDown`, `lightGrid`,
     * `mediumGray` -- which is a different vocabulary from the `ST_PresetPatternVal` a presentation
     * names a pattern by, and shares not one value with it. Nothing was written for them: the
     * pattern was dropped and the fill came out as its two colours and no pattern at all. Each now
     * stands for the preset it was asking for. `LIGHTTRELLIS` and `DARKTRELLIS` both answer
     * `trellis`, because the format has the one.
     */
    /** @deprecated 1.3.0 FILL_PATTERN_DARKDOWN said `darkDown`, which is a SpreadsheetML pattern; use FILL_PATTERN_DKDNDIAG */
    public const FILL_PATTERN_DARKDOWN = self::FILL_PATTERN_DKDNDIAG;
    /** @deprecated 1.3.0 FILL_PATTERN_DARKGRAY said `darkGray`, which is a SpreadsheetML pattern; use FILL_PATTERN_PCT75 */
    public const FILL_PATTERN_DARKGRAY = self::FILL_PATTERN_PCT75;
    /** @deprecated 1.3.0 FILL_PATTERN_DARKGRID said `darkGrid`, which is a SpreadsheetML pattern; use FILL_PATTERN_LGGRID */
    public const FILL_PATTERN_DARKGRID = self::FILL_PATTERN_LGGRID;
    /** @deprecated 1.3.0 FILL_PATTERN_DARKHORIZONTAL said `darkHorizontal`, which is a SpreadsheetML pattern; use FILL_PATTERN_DKHORZ */
    public const FILL_PATTERN_DARKHORIZONTAL = self::FILL_PATTERN_DKHORZ;
    /** @deprecated 1.3.0 FILL_PATTERN_DARKTRELLIS said `darkTrellis`, which is a SpreadsheetML pattern; use FILL_PATTERN_TRELLIS */
    public const FILL_PATTERN_DARKTRELLIS = self::FILL_PATTERN_TRELLIS;
    /** @deprecated 1.3.0 FILL_PATTERN_DARKUP said `darkUp`, which is a SpreadsheetML pattern; use FILL_PATTERN_DKUPDIAG */
    public const FILL_PATTERN_DARKUP = self::FILL_PATTERN_DKUPDIAG;
    /** @deprecated 1.3.0 FILL_PATTERN_DARKVERTICAL said `darkVertical`, which is a SpreadsheetML pattern; use FILL_PATTERN_DKVERT */
    public const FILL_PATTERN_DARKVERTICAL = self::FILL_PATTERN_DKVERT;
    /** @deprecated 1.3.0 FILL_PATTERN_GRAY0625 said `gray0625`, which is a SpreadsheetML pattern; use FILL_PATTERN_PCT5 */
    public const FILL_PATTERN_GRAY0625 = self::FILL_PATTERN_PCT5;
    /** @deprecated 1.3.0 FILL_PATTERN_GRAY125 said `gray125`, which is a SpreadsheetML pattern; use FILL_PATTERN_PCT10 */
    public const FILL_PATTERN_GRAY125 = self::FILL_PATTERN_PCT10;
    /** @deprecated 1.3.0 FILL_PATTERN_LIGHTDOWN said `lightDown`, which is a SpreadsheetML pattern; use FILL_PATTERN_LTDNDIAG */
    public const FILL_PATTERN_LIGHTDOWN = self::FILL_PATTERN_LTDNDIAG;
    /** @deprecated 1.3.0 FILL_PATTERN_LIGHTGRAY said `lightGray`, which is a SpreadsheetML pattern; use FILL_PATTERN_PCT25 */
    public const FILL_PATTERN_LIGHTGRAY = self::FILL_PATTERN_PCT25;
    /** @deprecated 1.3.0 FILL_PATTERN_LIGHTGRID said `lightGrid`, which is a SpreadsheetML pattern; use FILL_PATTERN_SMGRID */
    public const FILL_PATTERN_LIGHTGRID = self::FILL_PATTERN_SMGRID;
    /** @deprecated 1.3.0 FILL_PATTERN_LIGHTHORIZONTAL said `lightHorizontal`, which is a SpreadsheetML pattern; use FILL_PATTERN_LTHORZ */
    public const FILL_PATTERN_LIGHTHORIZONTAL = self::FILL_PATTERN_LTHORZ;
    /** @deprecated 1.3.0 FILL_PATTERN_LIGHTTRELLIS said `lightTrellis`, which is a SpreadsheetML pattern; use FILL_PATTERN_TRELLIS */
    public const FILL_PATTERN_LIGHTTRELLIS = self::FILL_PATTERN_TRELLIS;
    /** @deprecated 1.3.0 FILL_PATTERN_LIGHTUP said `lightUp`, which is a SpreadsheetML pattern; use FILL_PATTERN_LTUPDIAG */
    public const FILL_PATTERN_LIGHTUP = self::FILL_PATTERN_LTUPDIAG;
    /** @deprecated 1.3.0 FILL_PATTERN_LIGHTVERTICAL said `lightVertical`, which is a SpreadsheetML pattern; use FILL_PATTERN_LTVERT */
    public const FILL_PATTERN_LIGHTVERTICAL = self::FILL_PATTERN_LTVERT;
    /** @deprecated 1.3.0 FILL_PATTERN_MEDIUMGRAY said `mediumGray`, which is a SpreadsheetML pattern; use FILL_PATTERN_PCT50 */
    public const FILL_PATTERN_MEDIUMGRAY = self::FILL_PATTERN_PCT50;

    /**
     * Every preset pattern, to tell a pattern fill from a fill type that is none of the kinds
     * written before it.
     *
     * @var array<int, string>
     */
    public const PATTERN_TYPES = [
        self::FILL_PATTERN_PCT5, self::FILL_PATTERN_PCT10, self::FILL_PATTERN_PCT20, self::FILL_PATTERN_PCT25,
        self::FILL_PATTERN_PCT30, self::FILL_PATTERN_PCT40, self::FILL_PATTERN_PCT50, self::FILL_PATTERN_PCT60,
        self::FILL_PATTERN_PCT70, self::FILL_PATTERN_PCT75, self::FILL_PATTERN_PCT80, self::FILL_PATTERN_PCT90,
        self::FILL_PATTERN_HORZ, self::FILL_PATTERN_VERT, self::FILL_PATTERN_LTHORZ, self::FILL_PATTERN_LTVERT,
        self::FILL_PATTERN_DKHORZ, self::FILL_PATTERN_DKVERT, self::FILL_PATTERN_NARHORZ, self::FILL_PATTERN_NARVERT,
        self::FILL_PATTERN_DASHHORZ, self::FILL_PATTERN_DASHVERT, self::FILL_PATTERN_CROSS, self::FILL_PATTERN_DNDIAG,
        self::FILL_PATTERN_UPDIAG, self::FILL_PATTERN_LTDNDIAG, self::FILL_PATTERN_LTUPDIAG, self::FILL_PATTERN_DKDNDIAG,
        self::FILL_PATTERN_DKUPDIAG, self::FILL_PATTERN_WDDNDIAG, self::FILL_PATTERN_WDUPDIAG, self::FILL_PATTERN_DASHDNDIAG,
        self::FILL_PATTERN_DASHUPDIAG, self::FILL_PATTERN_DIAGCROSS, self::FILL_PATTERN_SMCHECK, self::FILL_PATTERN_LGCHECK,
        self::FILL_PATTERN_SMGRID, self::FILL_PATTERN_LGGRID, self::FILL_PATTERN_DOTGRID, self::FILL_PATTERN_SMCONFETTI,
        self::FILL_PATTERN_LGCONFETTI, self::FILL_PATTERN_HORZBRICK, self::FILL_PATTERN_DIAGBRICK, self::FILL_PATTERN_SOLIDDMND,
        self::FILL_PATTERN_OPENDMND, self::FILL_PATTERN_DOTDMND, self::FILL_PATTERN_PLAID, self::FILL_PATTERN_SPHERE,
        self::FILL_PATTERN_WEAVE, self::FILL_PATTERN_DIVOT, self::FILL_PATTERN_SHINGLE, self::FILL_PATTERN_WAVE,
        self::FILL_PATTERN_TRELLIS, self::FILL_PATTERN_ZIGZAG,
    ];

    /**
     * Fill type.
     *
     * @var string
     */
    private $fillType = self::FILL_NONE;

    /**
     * Rotation.
     *
     * @var float
     */
    private $rotation = 0.0;

    /**
     * Start color.
     *
     * @var Color
     */
    private $startColor;

    /**
     * End color.
     *
     * @var Color
     */
    private $endColor;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Fill.
     */
    public function __construct()
    {
        $this->startColor = new Color(Color::COLOR_BLACK);
        $this->endColor = new Color(Color::COLOR_WHITE);
    }

    /**
     * Get Fill Type.
     */
    public function getFillType(): string
    {
        return $this->fillType;
    }

    /**
     * Set Fill Type.
     *
     * @param string $pValue Fill type
     */
    public function setFillType(string $pValue = self::FILL_NONE): self
    {
        $this->fillType = $pValue;

        return $this;
    }

    /**
     * Get Rotation.
     */
    public function getRotation(): float
    {
        return $this->rotation;
    }

    /**
     * Set Rotation.
     */
    public function setRotation(float $pValue = 0): self
    {
        $this->rotation = $pValue;

        return $this;
    }

    /**
     * Get Start Color.
     */
    public function getStartColor(): Color
    {
        // It's a get but it may lead to a modified color which we won't detect but in which case we must bind.
        // So bind as an assurance.
        return $this->startColor;
    }

    /**
     * Set Start Color.
     */
    public function setStartColor(Color $pValue): self
    {
        $this->startColor = $pValue;

        return $this;
    }

    /**
     * Get End Color.
     */
    public function getEndColor(): Color
    {
        // It's a get but it may lead to a modified color which we won't detect but in which case we must bind.
        // So bind as an assurance.
        return $this->endColor;
    }

    /**
     * Set End Color.
     */
    public function setEndColor(Color $pValue): self
    {
        $this->endColor = $pValue;

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
            $this->getFillType()
            . $this->getRotation()
            . $this->getStartColor()->getHashCode()
            . $this->getEndColor()->getHashCode()
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
