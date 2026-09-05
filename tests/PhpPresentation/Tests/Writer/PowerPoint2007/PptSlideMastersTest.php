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

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use DK\OpenXml\OpenXmlPackage;
use PhpOffice\PhpPresentation\Shape\Drawing\File as ShapeDrawingFile;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\PptSlideMasters;
use PHPUnit\Framework\MockObject\MockBuilder;
use PHPUnit\Framework\TestCase;

/**
 * Test class for PowerPoint2007.
 *
 * @coversDefaultClass \PowerPoint2007
 */
class PptSlideMastersTest extends TestCase
{
    public function testWriteSlideMasterRelationships(): void
    {
        $writer = new PptSlideMasters();
        /** @var MockBuilder<SlideMaster> $mockBuilder */
        // @phpstan-ignore-next-line
        $mockBuilder = $this->getMockBuilder(SlideMaster::class);
        if (method_exists(get_class($mockBuilder), 'onlyMethods')) {
            /** @var \PHPUnit\Framework\MockObject\MockObject|SlideMaster $slideMaster */
            // @phpstan-ignore-next-line
            $slideMaster = $mockBuilder->onlyMethods(['getAllSlideLayouts', 'getRelsIndex', 'getShapeCollection'])
                ->getMock();
        } else {
            /** @var \PHPUnit\Framework\MockObject\MockObject|SlideMaster $slideMaster */
            // @phpstan-ignore-next-line
            $slideMaster = $mockBuilder->setMethods(['getAllSlideLayouts', 'getRelsIndex', 'getShapeCollection'])
                ->getMock();
        }

        $layouts = [new SlideLayout($slideMaster)];

        $slideMaster->expects(self::once())
            ->method('getAllSlideLayouts')
            ->willReturn($layouts);

        /** @var array<int, ShapeDrawingFile> $collection */
        $collection = [];
        // a drawing with no file of its own has no extension either, and the part it would be
        // written to could not be named
        foreach (range(1, 3) as $ignored) {
            $drawing = new ShapeDrawingFile();
            $drawing->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
            $collection[] = $drawing;
        }

        $slideMaster->expects(self::exactly(2))
            ->method('getShapeCollection')
            ->willReturn($collection);

        $source = '/ppt/slideMasters/slideMaster1.xml';
        $writer->setPackage(OpenXmlPackage::create());
        $writer->writeSlideMasterRelationships($source, $slideMaster);

        $relationships = $writer->getPackage()->getRelationships($source);

        self::assertCount(5, $relationships);

        $ids = [];
        foreach ($relationships as $relationship) {
            $ids[] = $relationship->getId();
        }
        self::assertSame(['rId1', 'rId2', 'rId3', 'rId4', 'rId5'], $ids);
    }
}
