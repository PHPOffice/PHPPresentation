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

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class PptPresentationTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender(): void
    {
        $this->assertZipFileExists('ppt/presentation.xml');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testDefaultTextLanguage(): void
    {
        $element = '/p:presentation/p:defaultTextStyle/a:defPPr/a:defRPr';
        $this->assertZipXmlAttributeEquals('ppt/presentation.xml', $element, 'lang', 'en-US');

        // text that names no language of its own is in the document's, not in French
        $this->oPresentation->getDocumentProperties()->setLanguage('uk-UA');
        $this->resetPresentationFile();
        $this->assertZipXmlAttributeEquals('ppt/presentation.xml', $element, 'lang', 'uk-UA');
        $this->assertIsSchemaECMA376Valid();
    }
}
