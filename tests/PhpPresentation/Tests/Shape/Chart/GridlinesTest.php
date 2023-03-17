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

namespace PhpOffice\PhpPresentation\Tests\Shape\Chart;

use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Style\Outline;
use PHPUnit\Framework\TestCase;

class GridlinesTest extends TestCase
{
    public function testConstruct(): void
    {
        $object = new Gridlines();

        $this->assertInstanceOf(Outline::class, $object->getOutline());
    }

    public function testGetSetOutline(): void
    {
        $object = new Gridlines();

        /** @var Outline $oStub */
        $oStub = $this->getMockBuilder(Outline::class)->getMock();

        $this->assertInstanceOf(Outline::class, $object->getOutline());
        $this->assertInstanceOf('PhpOffice\PhpPresentation\Shape\Chart\Gridlines', $object->setOutline($oStub));
        $this->assertInstanceOf(Outline::class, $object->getOutline());
    }
}
