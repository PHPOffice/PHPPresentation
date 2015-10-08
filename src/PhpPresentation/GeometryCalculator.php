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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation;

/**
 * PhpOffice\PhpPresentation\GeometryCalculator
 */
class GeometryCalculator
{
    const X = 'X';
    const Y = 'Y';

    /**
    * Calculate X and Y offsets for a set of shapes within a container such as a slide or group.
    *
    * @param  \PhpOffice\PhpPresentation\ShapeContainerInterface $container
    * @return array
    */
    public static function calculateOffsets(ShapeContainerInterface $container)
    {
        $offsets = array(self::X => 0, self::Y => 0);

        if ($container !== null && count($container->getShapeCollection()) != 0) {
            $shapes = $container->getShapeCollection();
            if ($shapes[0] !== null) {
                $offsets[self::X] = $shapes[0]->getOffsetX();
                $offsets[self::Y] = $shapes[0]->getOffsetY();
            }

            foreach ($shapes as $shape) {
                if ($shape !== null) {
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
    * @param  \PhpOffice\PhpPresentation\ShapeContainerInterface $container
    * @return array
    */
    public static function calculateExtents(ShapeContainerInterface $container)
    {
        $extents = array(self::X => 0, self::Y => 0);

        if ($container !== null && count($container->getShapeCollection()) != 0) {
            $shapes = $container->getShapeCollection();
            if ($shapes[0] !== null) {
                $extents[self::X] = $shapes[0]->getOffsetX() + $shapes[0]->getWidth();
                $extents[self::Y] = $shapes[0]->getOffsetY() + $shapes[0]->getHeight();
            }

            foreach ($shapes as $shape) {
                if ($shape !== null) {
                    $extentX = $shape->getOffsetX() + $shape->getWidth();
                    $extentY = $shape->getOffsetY() + $shape->getHeight();

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
