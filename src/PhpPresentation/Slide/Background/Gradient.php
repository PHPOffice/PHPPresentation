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
use PhpOffice\PhpPresentation\Style\Fill;

/**
 * A background painted with a gradient.
 *
 * The gradient is a fill, and it is the same fill a shape is painted with: the stops, the
 * angle and whether the gradient runs from the middle outwards are all the things
 * {@see Fill} holds. A background that is a gradient and a shape that is a gradient are
 * therefore one thing said in two places, which is why this holds one rather than a copy of
 * its parts.
 *
 * The fill is not optional, where the colour of a {@see Color} background is: a background
 * that names no colour is nothing at all, and a background that names no gradient is a
 * gradient of the two stops a fill holds by default.
 */
class Gradient extends AbstractBackground
{
    /**
     * @var Fill
     */
    protected $fill;

    public function __construct()
    {
        $this->fill = (new Fill())->setFillType(Fill::FILL_GRADIENT_LINEAR);
    }

    public function setFill(Fill $fill): self
    {
        $this->fill = $fill;

        return $this;
    }

    public function getFill(): Fill
    {
        return $this->fill;
    }
}
