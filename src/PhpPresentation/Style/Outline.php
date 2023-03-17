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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Style;

/**
 * \PhpOffice\PhpPresentation\Style\Outline.
 */
class Outline
{
    /**
     * @var Fill
     */
    protected $fill;

    /**
     * @var float
     */
    protected $width = 1;

    public function __construct()
    {
        $this->fill = new Fill();
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

    public function getWidth(): float
    {
        return $this->width;
    }

    /**
     * Value in points.
     */
    public function setWidth(float $pValue = 1): self
    {
        $this->width = $pValue;

        return $this;
    }
}
