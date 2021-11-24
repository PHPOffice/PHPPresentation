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

use PhpOffice\PhpPresentation\Slide\Background\SchemeColor;
use PhpOffice\PhpPresentation\Style\SchemeColor as StyleSchemeColor;
use PHPUnit\Framework\TestCase;

class SchemeColorTest extends TestCase
{
    public function testBasic(): void
    {
        $oStyle = new StyleSchemeColor();

        $object = new SchemeColor();

        $this->assertNull($object->getSchemeColor());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\SchemeColor', $object->setSchemeColor($oStyle));
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Style\\SchemeColor', $object->getSchemeColor());

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Slide\\Background\\SchemeColor', $object->setSchemeColor());
        $this->assertNull($object->getSchemeColor());
    }
}
