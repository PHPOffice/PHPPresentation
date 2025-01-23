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

namespace PhpOffice\PhpPresentation\Traits;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ShapeContainerInterface;

trait ShapeCollection
{
    /**
     * Collection of shapes.
     *
     * @var array<int, AbstractShape>
     */
    protected $shapeCollection = [];

    /**
     * Get collection of shapes.
     *
     * @return array<int, AbstractShape>
     */
    public function getShapeCollection(): array
    {
        return $this->shapeCollection;
    }

    /**
     * Search into collection of shapes for a name or/and a type.
     *
     * @return array<int, AbstractShape>
     */
    public function searchShapes(?string $name = null, ?string $type = null): array
    {
        $found = [];
        foreach ($this->getShapeCollection() as $shape) {
            if ($name && $shape->getName() !== $name) {
                continue;
            }
            if ($type && get_class($shape) !== $type) {
                continue;
            }

            $found[] = $shape;
        }

        return $found;
    }

    /**
     * Get collection of shapes.
     *
     * @param array<int, AbstractShape> $shapeCollection
     */
    public function setShapeCollection(array $shapeCollection = []): self
    {
        $this->shapeCollection = $shapeCollection;

        return $this;
    }

    /**
     * @return static
     */
    public function addShape(AbstractShape $shape)
    {
        if (!$shape->getContainer() && $this instanceof ShapeContainerInterface) {
            $shape->setContainer($this);
        } else {
            $this->shapeCollection[] = $shape;
        }

        return $this;
    }

    /**
     * @return static
     */
    public function unsetShape(int $key)
    {
        unset($this->shapeCollection[$key]);

        return $this;
    }
}
