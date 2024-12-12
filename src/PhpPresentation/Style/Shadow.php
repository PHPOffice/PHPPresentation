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
 * \PhpOffice\PhpPresentation\Style\Shadow.
 */
class Shadow implements ComparableInterface
{
    public const TYPE_SHADOW_INNER = 'innerShdw';
    public const TYPE_SHADOW_OUTER = 'outerShdw';
    public const TYPE_REFLECTION = 'reflection';

    // Shadow alignment
    public const SHADOW_BOTTOM = 'b';
    public const SHADOW_BOTTOM_LEFT = 'bl';
    public const SHADOW_BOTTOM_RIGHT = 'br';
    public const SHADOW_CENTER = 'ctr';
    public const SHADOW_LEFT = 'l';
    public const SHADOW_TOP = 't';
    public const SHADOW_TOP_LEFT = 'tl';
    public const SHADOW_TOP_RIGHT = 'tr';

    /**
     * Visible.
     *
     * @var bool
     */
    private $visible = false;

    /**
     * Blur radius.
     *
     * @var int
     */
    private $blurRadius = 6;

    /**
     * Shadow distance.
     *
     * @var int
     */
    private $distance = 2;

    /**
     * Shadow direction (in degrees).
     *
     * @var int
     */
    private $direction = 0;

    /**
     * Shadow alignment.
     *
     * @var string
     */
    private $alignment = self::SHADOW_BOTTOM_RIGHT;

    /**
     * @var null|Color
     */
    private $color;

    /**
     * @var int
     */
    private $alpha = 50;

    /**
     * @var string
     */
    private $type = self::TYPE_SHADOW_OUTER;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Shadow.
     */
    public function __construct()
    {
        $this->color = new Color(Color::COLOR_BLACK);
    }

    /**
     * Get Visible.
     */
    public function isVisible(): bool
    {
        return $this->visible;
    }

    /**
     * Set Visible.
     */
    public function setVisible(bool $pValue = false): self
    {
        $this->visible = $pValue;

        return $this;
    }

    /**
     * Get Blur radius.
     */
    public function getBlurRadius(): int
    {
        return $this->blurRadius;
    }

    /**
     * Set Blur radius.
     */
    public function setBlurRadius(int $pValue = 6): self
    {
        $this->blurRadius = $pValue;

        return $this;
    }

    /**
     * Get Shadow distance.
     */
    public function getDistance(): int
    {
        return $this->distance;
    }

    /**
     * Set Shadow distance.
     *
     * @return $this
     */
    public function setDistance(int $pValue = 2): self
    {
        $this->distance = $pValue;

        return $this;
    }

    /**
     * Get Shadow direction (in degrees).
     */
    public function getDirection(): int
    {
        return $this->direction;
    }

    /**
     * Set Shadow direction (in degrees).
     */
    public function setDirection(int $pValue = 0): self
    {
        $this->direction = $pValue;

        return $this;
    }

    /**
     * Get Shadow alignment.
     */
    public function getAlignment(): string
    {
        return $this->alignment;
    }

    /**
     * Set Shadow alignment.
     */
    public function setAlignment(string $pValue = self::SHADOW_BOTTOM_RIGHT): self
    {
        $this->alignment = $pValue;

        return $this;
    }

    /**
     * Get Color.
     */
    public function getColor(): ?Color
    {
        return $this->color;
    }

    /**
     * Set Color.
     */
    public function setColor(?Color $pValue = null): self
    {
        $this->color = $pValue;

        return $this;
    }

    /**
     * Get Alpha.
     */
    public function getAlpha(): int
    {
        return $this->alpha;
    }

    /**
     * Set Alpha.
     */
    public function setAlpha(int $pValue = 0): self
    {
        $this->alpha = $pValue;

        return $this;
    }

    /**
     * Get Type.
     */
    public function getType(): string
    {
        return $this->type;
    }

    /**
     * Set Type.
     */
    public function setType(string $pValue = self::TYPE_SHADOW_OUTER): self
    {
        if (!in_array(
            $pValue,
            [self::TYPE_REFLECTION, self::TYPE_SHADOW_INNER, self::TYPE_SHADOW_OUTER]
        )) {
            throw new NotAllowedValueException($pValue, [self::TYPE_REFLECTION, self::TYPE_SHADOW_INNER, self::TYPE_SHADOW_OUTER]);
        }

        $this->type = $pValue;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5(($this->visible ? 't' : 'f') . $this->blurRadius . $this->distance . $this->direction . $this->alignment . $this->type . $this->color->getHashCode() . $this->alpha . __CLASS__);
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
