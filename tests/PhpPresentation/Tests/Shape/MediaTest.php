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

namespace PhpPresentation\Tests\Shape;

use PhpOffice\PhpPresentation\Shape\Media;
use PHPUnit\Framework\TestCase;

class MediaTest extends TestCase
{
    public function testInheritance(): void
    {
        $object = new Media();

        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\Shape\\Drawing\\File', $object);
    }

    public function testMimeType(): void
    {
        $object = new Media();
        $object->setPath('file.mp4', false);
        $this->assertEquals('video/mp4', $object->getMimeType());
        $object->setPath('file.ogv', false);
        $this->assertEquals('video/ogg', $object->getMimeType());
        $object->setPath('file.wmv', false);
        $this->assertEquals('video/x-ms-wmv', $object->getMimeType());
        $object->setPath('file.xxx', false);
        $this->assertEquals('application/octet-stream', $object->getMimeType());
    }
}
