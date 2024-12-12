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

namespace PhpOffice\PhpPresentation\Shape\Chart;

use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;

class Marker
{
    public const SYMBOL_CIRCLE = 'circle';
    public const SYMBOL_DASH = 'dash';
    public const SYMBOL_DIAMOND = 'diamond';
    public const SYMBOL_DOT = 'dot';
    public const SYMBOL_NONE = 'none';
    public const SYMBOL_PLUS = 'plus';
    public const SYMBOL_SQUARE = 'square';
    public const SYMBOL_STAR = 'star';
    public const SYMBOL_TRIANGLE = 'triangle';
    public const SYMBOL_X = 'x';

    /**
     * @var array<int, string>
     */
    public static $arraySymbol = [
        self::SYMBOL_CIRCLE,
        self::SYMBOL_DASH,
        self::SYMBOL_DIAMOND,
        self::SYMBOL_DOT,
        self::SYMBOL_NONE,
        self::SYMBOL_PLUS,
        self::SYMBOL_SQUARE,
        self::SYMBOL_STAR,
        self::SYMBOL_TRIANGLE,
        self::SYMBOL_X,
    ];

    /**
     * @var string
     */
    protected $symbol = self::SYMBOL_NONE;

    /**
     * @var int
     */
    protected $size = 5;

    /**
     * @var Fill
     */
    protected $fill;

    /**
     * @var Border
     */
    protected $border;

    public function __construct()
    {
        $this->fill = new Fill();
        $this->border = new Border();
    }

    public function getSymbol(): string
    {
        return $this->symbol;
    }

    public function setSymbol(string $symbol = self::SYMBOL_NONE): self
    {
        $this->symbol = $symbol;

        return $this;
    }

    public function getSize(): int
    {
        return $this->size;
    }

    public function setSize(int $size = 5): self
    {
        $this->size = $size;

        return $this;
    }

    public function getFill(): Fill
    {
        return $this->fill;
    }

    public function setFill(Fill $fill): self
    {
        $this->fill = $fill;

        return $this;
    }

    public function getBorder(): Border
    {
        return $this->border;
    }

    public function setBorder(Border $border): self
    {
        $this->border = $border;

        return $this;
    }
}
