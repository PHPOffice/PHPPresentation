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

namespace PhpOffice\PhpPresentation\Tests;

use DOMDocument;
use DOMElement;
use DOMNode;
use DOMNodeList;
use DOMXPath;
use LibXMLError;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PHPUnit\Framework\TestCase;

class PhpPresentationTestCase extends TestCase
{
    /**
     * @var PhpPresentation|null
     */
    protected $oPresentation;

    /**
     * @var string
     */
    protected $filePath;

    /**
     * @var string
     */
    protected $workDirectory;

    /**
     * @var string
     */
    protected $writerName;

    /**
     * DOMDocument object.
     *
     * @var DOMDocument|null
     */
    private $xmlDom;

    /**
     * @var DOMXPath|null
     */
    private $xmlXPath;

    /**
     * File name.
     *
     * @var string|null
     */
    private $xmlFile;

    /**
     * @var bool
     */
    private $xmlInternalErrors;

    /**
     * @var bool
     */
    private $xmlDisableEntityLoader;

    /**
     * @var array<string, array<string, string>>
     */
    private $arrayOpenDocumentRNG = [
        '1.0' => [
            'META-INF/manifest.xml' => 'OpenDocument-manifest-schema-v1.0-os.rng',
            '*' => 'OpenDocument-strict-schema-v1.0-os.rng',
        ],
        '1.1' => [
            'META-INF/manifest.xml' => 'OpenDocument-manifest-schema-v1.1.rng',
            '*' => 'OpenDocument-strict-schema-v1.1.rng',
        ],
        '1.2' => [
            'META-INF/manifest.xml' => 'OpenDocument-v1.2-os-manifest-schema.rng',
            '*' => 'OpenDocument-v1.2-os-schema.rng',
        ],
    ];

    /**
     * Executed before each method of the class.
     */
    protected function setUp(): void
    {
        $this->xmlDisableEntityLoader = libxml_disable_entity_loader(false);
        $this->workDirectory = sys_get_temp_dir() . '/PhpPresentation_Unit_Test/';
        $this->oPresentation = new PhpPresentation();
        $this->filePath = tempnam(sys_get_temp_dir(), 'PhpPresentation');

        // Error XML
        libxml_clear_errors();
        $this->xmlInternalErrors = libxml_use_internal_errors(true);

        // Reset file
        $this->resetPresentationFile();
    }

    /**
     * Executed after each method of the class.
     */
    protected function tearDown(): void
    {
        libxml_disable_entity_loader($this->xmlDisableEntityLoader);
        libxml_use_internal_errors($this->xmlInternalErrors);
        $this->oPresentation = null;
        $this->resetPresentationFile();
    }

    /**
     * Delete directory.
     */
    private function deleteDir(string $dir): void
    {
        foreach (scandir($dir) as $file) {
            if ('.' === $file || '..' === $file) {
                continue;
            } elseif (is_file($dir . '/' . $file)) {
                unlink($dir . '/' . $file);
            } elseif (is_dir($dir . '/' . $file)) {
                $this->deleteDir($dir . '/' . $file);
            }
        }

        rmdir($dir);
    }

    protected function getXmlDom(string $file): DOMDocument
    {
        $baseFile = $file;
        if (null !== $this->xmlDom && $file === $this->xmlFile) {
            return $this->xmlDom;
        }

        $this->xmlXPath = null;
        $this->xmlFile = $file;

        $file = $this->workDirectory . '/' . $file;
        $this->xmlDom = new DOMDocument();
        $strContent = file_get_contents($file);
        // docProps/app.xml
        if ('docProps/app.xml' == $baseFile) {
            $strContent = str_replace(' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"', '', $strContent);
        }
        // docProps/custom.xml
        if ('docProps/custom.xml' == $baseFile) {
            $strContent = str_replace(' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"', '', $strContent);
        }
        // _rels/.rels
        if (false !== strpos($baseFile, '_rels/') && false !== strpos($baseFile, '.rels')) {
            $strContent = str_replace(' xmlns="http://schemas.openxmlformats.org/package/2006/relationships"', '', $strContent);
        }
        $this->xmlDom->loadXML($strContent);

        return $this->xmlDom;
    }

    /**
     * @return DOMNodeList<DOMNode>
     */
    private function getXmlNodeList(string $file, string $xpath): DOMNodeList
    {
        if (null === $this->xmlDom || $file !== $this->xmlFile) {
            $this->getXmlDom($file);
        }

        if (null === $this->xmlXPath) {
            $this->xmlXPath = new DOMXPath($this->xmlDom);
        }

        return $this->xmlXPath->query($xpath);
    }

    /**
     * @param string $writerName
     */
    protected function writePresentationFile(PhpPresentation $oPhpPresentation, $writerName): void
    {
        if (is_file($this->filePath)) {
            return;
        }

        $xmlWriter = IOFactory::createWriter($oPhpPresentation, $writerName);
        $xmlWriter->save($this->filePath);

        $zip = new \ZipArchive();
        $res = $zip->open($this->filePath);
        if (true === $res) {
            $zip->extractTo($this->workDirectory);
            $zip->close();
        }
    }

    protected function resetPresentationFile(): void
    {
        $this->xmlFile = null;
        $this->xmlDom = null;
        $this->xmlXPath = null;
        if (is_file($this->filePath)) {
            unlink($this->filePath);
        }
        if (is_dir($this->workDirectory)) {
            $this->deleteDir($this->workDirectory);
        }
        if (!is_dir($this->workDirectory)) {
            mkdir($this->workDirectory);
        }
    }

    /**
     * @param string $filePath
     */
    public function assertZipFileExists($filePath): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        self::assertTrue(is_file($this->workDirectory . $filePath));
    }

    /**
     * @param string $filePath
     */
    public function assertZipFileNotExists($filePath): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        self::assertFalse(is_file($this->workDirectory . $filePath));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     */
    public function assertZipXmlElementExists($filePath, $xPath): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertNotEquals(
            0,
            $nodeList->length,
            sprintf(
                'The element "%s" doesn\'t exist in the file "%s"',
                $xPath,
                $filePath
            )
        );
    }

    /**
     * @param string $filePath
     * @param string $xPath
     */
    public function assertZipXmlElementNotExists($filePath, $xPath): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertEquals(
            0,
            $nodeList->length,
            sprintf(
                'The element "%s" exist in the file "%s"',
                $xPath,
                $filePath
            )
        );
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param mixed $value
     */
    public function assertZipXmlElementEquals($filePath, $xPath, $value): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertEquals($nodeList->item(0)->nodeValue, $value);
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param int $num
     */
    public function assertZipXmlElementCount($filePath, $xPath, $num): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertEquals($nodeList->length, $num);
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     * @param mixed $value
     */
    public function assertZipXmlAttributeEquals($filePath, $xPath, $attribute, $value): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        /** @var DOMElement $nodeItem */
        $nodeItem = $nodeList->item(0);
        self::assertInstanceOf(DOMElement::class, $nodeItem);
        self::assertEquals($value, $nodeItem->getAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     * @param mixed $value
     */
    public function assertZipXmlAttributeStartsWith($filePath, $xPath, $attribute, $value): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        /** @var DOMElement $nodeItem */
        $nodeItem = $nodeList->item(0);
        self::assertInstanceOf(DOMElement::class, $nodeItem);
        self::assertStringStartsWith($value, $nodeItem->getAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     * @param mixed $value
     */
    public function assertZipXmlAttributeEndsWith($filePath, $xPath, $attribute, $value): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        /** @var DOMElement $nodeItem */
        $nodeItem = $nodeList->item(0);
        self::assertInstanceOf(DOMElement::class, $nodeItem);
        self::assertStringEndsWith($value, $nodeItem->getAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     * @param mixed $value
     */
    public function assertZipXmlAttributeContains($filePath, $xPath, $attribute, $value): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        /** @var DOMElement $nodeItem */
        $nodeItem = $nodeList->item(0);
        self::assertInstanceOf(DOMElement::class, $nodeItem);
        self::assertStringContainsString($value, $nodeItem->getAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     */
    public function assertZipXmlAttributeExists($filePath, $xPath, $attribute): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        /** @var DOMElement $nodeItem */
        $nodeItem = $nodeList->item(0);
        self::assertInstanceOf(DOMElement::class, $nodeItem);
        self::assertTrue($nodeItem->hasAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     */
    public function assertZipXmlAttributeNotExists($filePath, $xPath, $attribute): void
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        /** @var DOMElement $nodeItem */
        $nodeItem = $nodeList->item(0);
        self::assertInstanceOf(DOMElement::class, $nodeItem);
        self::assertFalse($nodeItem->hasAttribute($attribute));
    }

    public function assertIsSchemaECMA376Valid(): void
    {
        // validate all XML files
        $path = realpath($this->workDirectory . '/ppt');
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($path));

        foreach ($iterator as $file) {
            /** @var \SplFileInfo $file */
            if ('xml' !== $file->getExtension()) {
                continue;
            }

            $fileName = str_replace('\\', '/', substr($file->getRealPath(), strlen($path) + 1));
            $dom = $this->getXmlDom('ppt/' . $fileName);
            $xmlSource = $dom->saveXML();

            $dom->loadXML($xmlSource);
            $dom->schemaValidate(__DIR__ . '/../../../resources/schema/ecma-376/pml.xsd');

            $error = libxml_get_last_error();
            if ($error instanceof LibXMLError) {
                $this->failXmlError($error, $fileName, $xmlSource);
            }
        }
        unset($iterator);
    }

    public function assertIsSchemaOOXMLValid(): void
    {
        // validate all XML files
        $path = realpath($this->workDirectory . '/ppt');
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($path));

        foreach ($iterator as $file) {
            /** @var \SplFileInfo $file */
            if ('xml' !== $file->getExtension()) {
                continue;
            }

            $fileName = str_replace('\\', '/', substr($file->getRealPath(), strlen($path) + 1));
            $dom = $this->getXmlDom('ppt/' . $fileName);
            $xmlSource = $dom->saveXML();
            // In the ISO/ECMA standard the namespace has changed from
            // http://schemas.openxmlformats.org/ to http://purl.oclc.org/ooxml/
            // We need to use the http://purl.oclc.org/ooxml/ namespace to validate
            // the xml against the current schema
            $xmlSource = str_replace([
                'http://schemas.openxmlformats.org/drawingml/2006/main',
                'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'http://schemas.openxmlformats.org/presentationml/2006/main',
            ], [
                'http://purl.oclc.org/ooxml/drawingml/main',
                'http://purl.oclc.org/ooxml/drawingml/chart',
                'http://purl.oclc.org/ooxml/officeDocument/relationships',
                'http://purl.oclc.org/ooxml/presentationml/main',
            ], $xmlSource);

            $dom->loadXML($xmlSource);
            $dom->schemaValidate(__DIR__ . '/../../../resources/schema/ooxml/pml.xsd');

            $error = libxml_get_last_error();
            if ($error instanceof LibXMLError) {
                $this->failXmlError($error, $fileName, $xmlSource);
            }
        }
        unset($iterator);
    }

    public function assertIsSchemaOpenDocumentValid(string $version = '1.0', bool $triggerError = true): bool
    {
        if (!array_key_exists($version, $this->arrayOpenDocumentRNG)) {
            self::fail('assertIsSchemaOpenDocumentValid > Use a valid version');
        }

        // validate all XML files
        $path = realpath($this->workDirectory);
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($path));

        $isValid = true;
        foreach ($iterator as $file) {
            /** @var \SplFileInfo $file */
            if ('xml' !== $file->getExtension()) {
                continue;
            }

            $fileName = str_replace('\\', '/', substr($file->getRealPath(), strlen($path) + 1));
            $dom = $this->getXmlDom($fileName);
            $xmlSource = $dom->saveXML();

            $dom->loadXML($xmlSource);
            $pathRNG = __DIR__ . '/../../../resources/schema/opendocument/' . $version . '/';
            if (isset($this->arrayOpenDocumentRNG[$version][$fileName])) {
                $pathRNG .= $this->arrayOpenDocumentRNG[$version][$fileName];
            } else {
                $pathRNG .= $this->arrayOpenDocumentRNG[$version]['*'];
            }
            $dom->relaxNGValidate($pathRNG);

            $error = libxml_get_last_error();
            if ($error instanceof LibXMLError) {
                if ($triggerError) {
                    $this->failXmlError($error, $fileName, $xmlSource, ['version' => $version]);
                }
                $isValid = false;
            }
        }
        unset($iterator);

        return $isValid;
    }

    public function assertIsSchemaOpenDocumentNotValid(string $version = '1.0'): void
    {
        $isValid = $this->assertIsSchemaOpenDocumentValid($version, false);
        if ($isValid) {
            self::fail('Failed : This document is currently valid (Schema version: ' . $version . ')');
        }
    }

    /**
     * @param array<string, string> $params
     */
    protected function failXmlError(LibXMLError $error, string $fileName, string $source, array $params = []): void
    {
        switch ($error->level) {
            case LIBXML_ERR_WARNING:
                $errorType = 'warning';
                break;
            case LIBXML_ERR_ERROR:
                $errorType = 'error';
                break;
            case LIBXML_ERR_FATAL:
                $errorType = 'fatal';
                break;
            default:
                $errorType = 'Error';
                break;
        }
        $errorLine = (int) $error->line;
        $contents = explode("\n", $source);
        $lines = [];
        if (isset($contents[$errorLine - 2])) {
            $lines[] = '>> ' . $contents[$errorLine - 2];
        }
        if (isset($contents[$errorLine - 1])) {
            $lines[] = '>>> ' . $contents[$errorLine - 1];
        }
        if (isset($contents[$errorLine])) {
            $lines[] = '>> ' . $contents[$errorLine];
        }
        $paramStr = '';
        if (!empty($params)) {
            $paramStr .= "\n" . ' - Parameters :' . "\n";
            foreach ($params as $key => $val) {
                $paramStr .= '   - ' . $key . ' : ' . $val . "\n";
            }
        }
        self::fail(sprintf(
            "Validation %s :\n - File : %s\n - Line : %s\n - Message : %s - Lines :\n%s%s",
            $errorType,
            $fileName,
            $error->line,
            $error->message,
            implode(PHP_EOL, $lines),
            $paramStr
        ));
    }
}
