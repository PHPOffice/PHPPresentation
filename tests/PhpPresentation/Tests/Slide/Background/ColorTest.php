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

namespace PhpOffice\PhpPresentation\Tests\Slide\Background;

use PhpOffice\PhpPresentation\Slide\Background\Color;
use PhpOffice\PhpPresentation\Style\Color as StyleColor;
use PHPUnit\Framework\TestCase;

class ColorTest extends TestCase
{
    public function testColor(): void
    {
        $object = new Color();

        $oStyleColor = new StyleColor();
        $oStyleColor->setRGB('123456');

        $this->assertNull($object->getColor());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Color', $object->setColor($oStyleColor));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\Color', $object->getColor());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\Color', $object->setColor());
        $this->assertNull($object->getColor());
    }
}
