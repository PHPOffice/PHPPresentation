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

abstract class AbstractTypeLine extends AbstractType
{
    /**
     * Is Line Smooth?
     *
     * @var bool
     */
    protected $isSmooth = false;

    /**
     * Is Line Smooth?
     */
    public function isSmooth(): bool
    {
        return $this->isSmooth;
    }

    /**
     * Set Line Smoothness.
     */
    public function setIsSmooth(bool $value = true): self
    {
        $this->isSmooth = $value;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5($this->isSmooth() ? '1' : '0');
    }
}
