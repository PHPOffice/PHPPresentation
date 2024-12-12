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

namespace PhpOffice\PhpPresentation\Slide\Background;

use PhpOffice\PhpPresentation\Slide\AbstractBackground;
use PhpOffice\PhpPresentation\Style\SchemeColor as StyleSchemeColor;

class SchemeColor extends AbstractBackground
{
    /**
     * @var null|StyleSchemeColor
     */
    protected $schemeColor;

    /**
     * @return $this
     */
    public function setSchemeColor(?StyleSchemeColor $color = null): self
    {
        $this->schemeColor = $color;

        return $this;
    }

    public function getSchemeColor(): ?StyleSchemeColor
    {
        return $this->schemeColor;
    }
}
