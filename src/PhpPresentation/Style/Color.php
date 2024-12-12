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
 * \PhpOffice\PhpPresentation\Style\Color.
 */
class Color implements ComparableInterface
{
    // Colors
    public const COLOR_BLACK = 'FF000000';
    public const COLOR_WHITE = 'FFFFFFFF';
    public const COLOR_RED = 'FFFF0000';
    public const COLOR_DARKRED = 'FF800000';
    public const COLOR_BLUE = 'FF0000FF';
    public const COLOR_DARKBLUE = 'FF000080';
    public const COLOR_GREEN = 'FF00FF00';
    public const COLOR_DARKGREEN = 'FF008000';
    public const COLOR_YELLOW = 'FFFFFF00';
    public const COLOR_DARKYELLOW = 'FF808000';

    /**
     * ARGB - Alpha RGB.
     *
     * @var string
     */
    private $argb;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Color.
     *
     * @param string $pARGB
     */
    public function __construct($pARGB = self::COLOR_BLACK)
    {
        // Initialise values
        $this->argb = $pARGB;
    }

    /**
     * Get ARGB.
     *
     * @return string
     */
    public function getARGB()
    {
        return $this->argb;
    }

    /**
     * Set ARGB.
     *
     * @param string $pValue
     *
     * @return Color
     */
    public function setARGB($pValue = self::COLOR_BLACK)
    {
        if ('' == $pValue) {
            $pValue = self::COLOR_BLACK;
        }
        $this->argb = $pValue;

        return $this;
    }

    /**
     * Get the alpha % of the ARGB
     * Will return 100 if no ARGB.
     */
    public function getAlpha(): int
    {
        $alpha = 100;
        if (strlen($this->argb) >= 6) {
            $dec = hexdec(substr($this->argb, 0, 2));
            $alpha = (int) number_format(($dec / 255) * 100, 0);
        }

        return $alpha;
    }

    /**
     * Set the alpha % of the ARGB.
     *
     * @return $this
     */
    public function setAlpha(int $alpha = 100): self
    {
        if ($alpha < 0) {
            $alpha = 0;
        }
        if ($alpha > 100) {
            $alpha = 100;
        }
        $alpha = round(($alpha / 100) * 255);
        $alpha = dechex((int) $alpha);
        $alpha = str_pad($alpha, 2, '0', STR_PAD_LEFT);
        $this->argb = $alpha . substr($this->argb, 2);

        return $this;
    }

    /**
     * Get RGB.
     *
     * @return string
     */
    public function getRGB()
    {
        if (6 == strlen($this->argb)) {
            return $this->argb;
        }

        return substr($this->argb, 2);
    }

    /**
     * Set RGB.
     *
     * @param string $pValue
     * @param string $pAlpha
     *
     * @return Color
     */
    public function setRGB($pValue = '000000', $pAlpha = 'FF')
    {
        if ('' == $pValue) {
            $pValue = '000000';
        }
        if ('' == $pAlpha) {
            $pAlpha = 'FF';
        }
        $this->argb = $pAlpha . $pValue;

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
            $this->argb
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
