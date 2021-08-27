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

use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class DocPropsCoreTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender(): void
    {
        $this->assertZipFileExists('docProps/core.xml');
        $this->assertZipXmlElementNotExists('docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testDocumentProperties(): void
    {
        $expected = 'aAbBcDeE';

        $this->oPresentation->getDocumentProperties()
            ->setCreator($expected)
            ->setTitle($expected)
            ->setDescription($expected)
            ->setSubject($expected)
            ->setKeywords($expected)
            ->setCategory($expected);

        $this->assertZipFileExists('docProps/core.xml');
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/dc:creator');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/dc:creator', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/dc:title');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/dc:title', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/dc:description');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/dc:description', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/dc:subject');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/dc:subject', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/cp:keywords');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/cp:keywords', $expected);
        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/cp:category');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/cp:category', $expected);
        $this->assertIsSchemaECMA376Valid();
    }

    public function testMarkAsFinalTrue(): void
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(true);

        $this->assertZipXmlElementExists('docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
        $this->assertZipXmlElementEquals('docProps/core.xml', '/cp:coreProperties/cp:contentStatus', 'Final');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testMarkAsFinalFalse(): void
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(false);

        $this->assertZipXmlElementNotExists('docProps/core.xml', '/cp:coreProperties/cp:contentStatus');
        $this->assertIsSchemaECMA376Valid();
    }
}
