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

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Type\Bar.
 */
abstract class AbstractTypePie extends AbstractType
{
    /**
     * Create a new self instance.
     */
    public function __construct()
    {
        $this->hasAxisX = false;
        $this->hasAxisY = false;
    }

    /**
     * Explosion of the Pie.
     *
     * @var int
     */
    protected $explosion = 0;

    /**
     * Set explosion.
     */
    public function setExplosion(int $value = 0): self
    {
        $this->explosion = $value;

        return $this;
    }

    /**
     * Get orientation.
     */
    public function getExplosion(): int
    {
        return $this->explosion;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        $hash = '';
        foreach ($this->getSeries() as $series) {
            $hash .= $series->getHashCode();
        }

        return $hash;
    }
}
