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

class PptViewPropsTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender(): void
    {
        $expectedElement = '/p:viewPr';

        $this->assertZipFileExists('ppt/viewProps.xml');
        $this->assertZipXmlElementExists('ppt/viewProps.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', $expectedElement, 'showComments', 0);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', $expectedElement, 'lastView', PresentationProperties::VIEW_SLIDE);
        $this->assertIsSchemaECMA376Valid();
    }

    public function testCommentVisible(): void
    {
        $expectedElement = '/p:viewPr';

        $this->oPresentation->getPresentationProperties()->setCommentVisible(true);

        $this->assertZipFileExists('ppt/viewProps.xml');
        $this->assertZipXmlElementExists('ppt/viewProps.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', $expectedElement, 'showComments', 1);
        $this->assertIsSchemaECMA376Valid();
    }

    public function testLastView(): void
    {
        $expectedElement = '/p:viewPr';
        $expectedLastView = PresentationProperties::VIEW_OUTLINE;

        $this->oPresentation->getPresentationProperties()->setLastView($expectedLastView);

        $this->assertZipFileExists('ppt/viewProps.xml');
        $this->assertZipXmlElementExists('ppt/viewProps.xml', $expectedElement);
        $this->assertZipXmlAttributeEquals('ppt/viewProps.xml', $expectedElement, 'lastView', $expectedLastView);
        $this->assertIsSchemaECMA376Valid();
    }
}
