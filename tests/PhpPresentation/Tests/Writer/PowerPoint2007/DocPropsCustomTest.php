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

use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

class DocPropsCustomTest extends PhpPresentationTestCase
{
    protected $writerName = 'PowerPoint2007';

    public function testRender(): void
    {
        $this->assertZipFileExists('docProps/custom.xml');
        $this->assertZipXmlElementNotExists('docProps/custom.xml', '/Properties/property[@name="_MarkAsFinal"]');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testMarkAsFinalTrue(): void
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(true);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="_MarkAsFinal"]/vt:bool');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testMarkAsFinalFalse(): void
    {
        $this->oPresentation->getPresentationProperties()->markAsFinal(false);

        $this->assertZipXmlElementNotExists('docProps/custom.xml', '/Properties/property[@name="_MarkAsFinal"]');
        $this->assertIsSchemaECMA376Valid();
    }

    public function testCustomPropertiesBoolean(): void
    {
        $this->oPresentation->getDocumentProperties()->setCustomProperty('pName', false, null);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:bool');
        $this->assertZipXmlElementEquals('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:bool', 'false');
    }

    public function testCustomPropertiesDate(): void
    {
        $value = time();
        $this->oPresentation->getDocumentProperties()->setCustomProperty('pName', $value, DocumentProperties::PROPERTY_TYPE_DATE);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:filetime');
        $this->assertZipXmlElementEquals('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:filetime', date(DATE_W3C, $value));
    }

    public function testCustomPropertiesFloat(): void
    {
        $this->oPresentation->getDocumentProperties()->setCustomProperty('pName', 2.1, null);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:r8');
        $this->assertZipXmlElementEquals('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:r8', 2.1);
    }

    public function testCustomPropertiesInteger(): void
    {
        $this->oPresentation->getDocumentProperties()->setCustomProperty('pName', 2, null);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:i4');
        $this->assertZipXmlElementEquals('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:i4', 2);
    }

    public function testCustomPropertiesNull(): void
    {
        $this->oPresentation->getDocumentProperties()->setCustomProperty('pName', null, null);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:lpwstr');
        $this->assertZipXmlElementEquals('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:lpwstr', '');
    }

    public function testCustomPropertiesString(): void
    {
        $this->oPresentation->getDocumentProperties()->setCustomProperty('pName', 'pValue', null);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:lpwstr');
        $this->assertZipXmlElementEquals('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:lpwstr', 'pValue');
    }

    public function testCustomPropertiesUnknown(): void
    {
        $value = time();
        $this->oPresentation->getDocumentProperties()->setCustomProperty('pName', (string) $value, DocumentProperties::PROPERTY_TYPE_UNKNOWN);

        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]');
        $this->assertZipXmlElementExists('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:lpwstr');
        $this->assertZipXmlElementEquals('docProps/custom.xml', '/Properties/property[@pid="2"][@fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"][@name="pName"]/vt:lpwstr', $value);
    }
}
