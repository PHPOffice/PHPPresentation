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

namespace PhpOffice\PhpPresentation\Shape\Chart\Type;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * self.
 */
class Doughnut extends AbstractTypePie implements ComparableInterface
{
    /**
     * Hole Size.
     *
     * @var int
     */
    protected $holeSize = 50;

    /**
     * Chart Direction & Rotation
     *
     * @var string
     */
    public const DIR_CLOCKWISE        = 'clockWise';
    public const DIR_COUNTERCLOCKWISE = 'counterClockwise';
    private ?int $firstSliceAngle = null;   // 0-359
    private string $pieDirection = self::DIR_CLOCKWISE;

    /**
     * @return int
     */
    public function getHoleSize()
    {
        return $this->holeSize;
    }

    /**
     * @param int $holeSize
     *
     * @return Doughnut
     *
     * @see https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.charts.holesize(v=office.14).aspx
     */
    public function setHoleSize($holeSize = 50)
    {
        if ($holeSize < 10) {
            $holeSize = 10;
        }
        if ($holeSize > 90) {
            $holeSize = 90;
        }
        $this->holeSize = $holeSize;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5(parent::getHashCode() . __CLASS__);
    }

    public function setFirstSliceAngle(int $angle): self
    {
        $this->firstSliceAngle = (($angle % 360) + 360) % 360;
        return $this;
    }

    public function getFirstSliceAngle(): ?int
    {
        return $this->firstSliceAngle;
    }

    public function setPieDirection(string $dir): self
    {
        if (!in_array($dir, [self::DIR_CLOCKWISE, self::DIR_COUNTERCLOCKWISE], true)) {
            throw new \InvalidArgumentException('Invalid pie direction');
        }
        $this->pieDirection = $dir;

        return $this;
    }

    public function getPieDirection(): string
    {
        return $this->pieDirection;
    }
}
