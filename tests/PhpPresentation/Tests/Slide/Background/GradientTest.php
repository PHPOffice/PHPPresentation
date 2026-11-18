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

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Slide\Background\Gradient;
use PhpOffice\PhpPresentation\Style\Fill;
use PHPUnit\Framework\TestCase;

class GradientTest extends TestCase
{
    public function testFill(): void
    {
        $object = new Gradient();

        // A background that names no gradient is a linear one of the two stops a fill holds
        self::assertEquals(Fill::FILL_GRADIENT_LINEAR, $object->getFill()->getFillType());

        $fill = (new Fill())->setFillType(Fill::FILL_GRADIENT_PATH);
        self::assertInstanceOf(Gradient::class, $object->setFill($fill));
        self::assertSame($fill, $object->getFill());
    }
}
