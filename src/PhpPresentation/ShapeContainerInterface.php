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

namespace PhpOffice\PhpPresentation;

use ArrayObject;

/**
 * PhpOffice\PhpPresentation\ShapeContainerInterface.
 */
interface ShapeContainerInterface
{
    /**
     * Get collection of shapes.
     *
     * @return array<int, AbstractShape>|ArrayObject<int, AbstractShape>
     */
    public function getShapeCollection();

    /**
     * Add shape to slide.
     *
     * @return static
     */
    public function addShape(AbstractShape $shape);

    /**
     * Unset shape from the collection.
     *
     * @return static
     */
    public function unsetShape(int $key);

    /**
     * Get X Offset.
     */
    public function getOffsetX(): int;

    /**
     * Get Y Offset.
     */
    public function getOffsetY(): int;

    /**
     * Get X Extent.
     */
    public function getExtentX(): int;

    /**
     * Get Y Extent.
     */
    public function getExtentY(): int;

    public function getHashCode(): string;
}
