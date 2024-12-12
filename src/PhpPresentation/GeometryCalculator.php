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

/**
 * PhpOffice\PhpPresentation\GeometryCalculator.
 */
class GeometryCalculator
{
    public const X = 'X';
    public const Y = 'Y';

    /**
     * Calculate X and Y offsets for a set of shapes within a container such as a slide or group.
     *
     * @return array<string, int>
     */
    public static function calculateOffsets(ShapeContainerInterface $container): array
    {
        $offsets = [self::X => 0, self::Y => 0];

        if (null !== $container && 0 != count($container->getShapeCollection())) {
            $shapes = $container->getShapeCollection();
            if (null !== $shapes[0]) {
                $offsets[self::X] = $shapes[0]->getOffsetX();
                $offsets[self::Y] = $shapes[0]->getOffsetY();
            }

            foreach ($shapes as $shape) {
                if (null !== $shape) {
                    if ($shape->getOffsetX() < $offsets[self::X]) {
                        $offsets[self::X] = $shape->getOffsetX();
                    }

                    if ($shape->getOffsetY() < $offsets[self::Y]) {
                        $offsets[self::Y] = $shape->getOffsetY();
                    }
                }
            }
        }

        return $offsets;
    }

    /**
     * Calculate X and Y extents for a set of shapes within a container such as a slide or group.
     *
     * @return array<string, int>
     */
    public static function calculateExtents(ShapeContainerInterface $container): array
    {
        /** @var array<string, int> $extents */
        $extents = [self::X => 0, self::Y => 0];

        if (null !== $container && 0 != count($container->getShapeCollection())) {
            $shapes = $container->getShapeCollection();
            if (null !== $shapes[0]) {
                $extents[self::X] = (int) ($shapes[0]->getOffsetX() + $shapes[0]->getWidth());
                $extents[self::Y] = (int) ($shapes[0]->getOffsetY() + $shapes[0]->getHeight());
            }

            foreach ($shapes as $shape) {
                if (null !== $shape) {
                    $extentX = (int) ($shape->getOffsetX() + $shape->getWidth());
                    $extentY = (int) ($shape->getOffsetY() + $shape->getHeight());

                    if ($extentX > $extents[self::X]) {
                        $extents[self::X] = $extentX;
                    }

                    if ($extentY > $extents[self::Y]) {
                        $extents[self::Y] = $extentY;
                    }
                }
            }
        }

        return $extents;
    }
}
