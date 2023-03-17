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

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptPresPropsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender(): void
    {
        $this->assertZipFileExists('ppt/presProps.xml');
        $element = '/p:presentationPr/p:extLst/p:ext';
        $this->assertZipXmlElementExists('ppt/presProps.xml', $element);
        $this->assertZipXmlAttributeEquals('ppt/presProps.xml', $element, 'uri', '{E76CE94A-603C-4142-B9EB-6D1370010A27}');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testLoopContinuously(): void
    {
        $this->assertZipFileExists('ppt/presProps.xml');
        $element = '/p:presentationPr/p:showPr';
        $this->assertZipXmlElementExists('ppt/presProps.xml', $element);
        $this->assertZipXmlAttributeNotExists('ppt/presProps.xml', $element, 'loop');
        $this->assertIsSchemaECMA376Valid();

        $this->oPresentation->getPresentationProperties()->setLoopContinuouslyUntilEsc(true);
        $this->resetPresentationFile();

        $this->assertZipFileExists('ppt/presProps.xml');
        $element = '/p:presentationPr/p:showPr';
        $this->assertZipXmlElementExists('ppt/presProps.xml', $element);
        $this->assertZipXmlAttributeExists('ppt/presProps.xml', $element, 'loop');
        $this->assertZipXmlAttributeEquals('ppt/presProps.xml', $element, 'loop', 1);
        $this->assertIsSchemaECMA376Valid();
    }

    /**
     * @dataProvider dataProviderShowType
     */
    public function testShowType(string $slideshowType, string $element): void
    {
        $this->oPresentation->getPresentationProperties()->setSlideshowType($slideshowType);

        $this->assertZipFileExists('ppt/presProps.xml');
        $this->assertZipXmlElementExists('ppt/presProps.xml', $element);
        $this->assertIsSchemaECMA376Valid();
    }

    /**
     * @return array<array<string>>
     */
    public function dataProviderShowType(): array
    {
        return [
            [
                PresentationProperties::SLIDESHOW_TYPE_PRESENT,
                '/p:presentationPr/p:showPr/p:present',
            ],
            [
                PresentationProperties::SLIDESHOW_TYPE_BROWSE,
                '/p:presentationPr/p:showPr/p:browse',
            ],
            [
                PresentationProperties::SLIDESHOW_TYPE_KIOSK,
                '/p:presentationPr/p:showPr/p:kiosk',
            ],
        ];
    }
}
