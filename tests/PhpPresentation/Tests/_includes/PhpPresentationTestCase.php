<?php

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PHPUnit\Framework\TestCase;

class PhpPresentationTestCase extends TestCase
{
    /**
     * @var PhpPresentation
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
     * DOMDocument object
     *
     * @var \DOMDocument
     */
    private $xmlDom;

    /**
     * DOMXpath object
     *
     * @var \DOMXpath
     */
    private $xmlXPath;

    /**
     * File name
     *
     * @var string
     */
    private $xmlFile;

    /**
     * @var boolean
     */
    private $xmlInternalErrors;

    /**
     * @var boolean
     */
    private $xmlDisableEntityLoader;

    /**
     * @var array
     */
    private $arrayOpenDocumentRNG = array(
        '1.0' => array(
            'META-INF/manifest.xml' => 'OpenDocument-manifest-schema-v1.0-os.rng',
            '*' => 'OpenDocument-strict-schema-v1.0-os.rng',
        ),
        '1.1' => array(
            'META-INF/manifest.xml' => 'OpenDocument-manifest-schema-v1.1.rng',
            '*' => 'OpenDocument-strict-schema-v1.1.rng',
        ),
        '1.2' => array(
            'META-INF/manifest.xml' => 'OpenDocument-v1.2-os-manifest-schema.rng',
            '*' => 'OpenDocument-v1.2-os-schema.rng',
        )
    );

    /**
     * Executed before each method of the class
     */
    protected function setUp()
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
     * Executed after each method of the class
     */
    protected function tearDown()
    {
        libxml_disable_entity_loader($this->xmlDisableEntityLoader);
        libxml_use_internal_errors($this->xmlInternalErrors);
        $this->oPresentation = null;
        $this->resetPresentationFile();
    }

    /**
     * Delete directory
     *
     * @param string $dir
     */
    private function deleteDir($dir)
    {
        foreach (scandir($dir) as $file) {
            if ($file === '.' || $file === '..') {
                continue;
            } elseif (is_file($dir . "/" . $file)) {
                unlink($dir . "/" . $file);
            } elseif (is_dir($dir . "/" . $file)) {
                $this->deleteDir($dir . "/" . $file);
            }
        }

        rmdir($dir);
    }

    protected function getXmlDom($file)
    {
        $baseFile = $file;
        if (null !== $this->xmlDom && $file === $this->xmlFile) {
            return $this->xmlDom;
        }

        $this->xmlXPath = null;
        $this->xmlFile = $file;

        $file = $this->workDirectory . '/' . $file;
        $this->xmlDom = new \DOMDocument();
        $strContent = file_get_contents($file);
        // docProps/app.xml
        if ($baseFile == 'docProps/app.xml') {
            $strContent = str_replace(' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"', '', $strContent);
        }
        // docProps/custom.xml
        if ($baseFile == 'docProps/custom.xml') {
            $strContent = str_replace(' xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"', '', $strContent);
        }
        // _rels/.rels
        if (strpos($baseFile, '_rels/') !== false && strpos($baseFile, '.rels') !== false) {
            $strContent = str_replace(' xmlns="http://schemas.openxmlformats.org/package/2006/relationships"', '', $strContent);
        }
        $this->xmlDom->loadXML($strContent);
        return $this->xmlDom;
    }

    private function getXmlNodeList($file, $xpath)
    {
        if ($this->xmlDom === null || $file !== $this->xmlFile) {
            $this->getXmlDom($file);
        }

        if (null === $this->xmlXPath) {
            $this->xmlXPath = new \DOMXpath($this->xmlDom);
        }

        return $this->xmlXPath->query($xpath);
    }

    /**
     * @param PhpPresentation $oPhpPresentation
     * @param string $writerName
     */
    protected function writePresentationFile(PhpPresentation $oPhpPresentation, $writerName)
    {
        if (is_file($this->filePath)) {
            return;
        }

        $xmlWriter = IOFactory::createWriter($oPhpPresentation, $writerName);
        $xmlWriter->save($this->filePath);

        $zip = new \ZipArchive;
        $res = $zip->open($this->filePath);
        if ($res === true) {
            $zip->extractTo($this->workDirectory);
            $zip->close();
        }
    }

    protected function resetPresentationFile()
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
    public function assertZipFileExists($filePath)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        self::assertTrue(is_file($this->workDirectory . $filePath));
    }

    /**
     * @param string $filePath
     */
    public function assertZipFileNotExists($filePath)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        self::assertFalse(is_file($this->workDirectory . $filePath));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     */
    public function assertZipXmlElementExists($filePath, $xPath)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertNotEquals(0, $nodeList->length);
    }

    /**
     * @param string $filePath
     * @param string $xPath
     */
    public function assertZipXmlElementNotExists($filePath, $xPath)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertEquals(0, $nodeList->length);
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param mixed $value
     */
    public function assertZipXmlElementEquals($filePath, $xPath, $value)
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
    public function assertZipXmlElementCount($filePath, $xPath, $num)
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
    public function assertZipXmlAttributeEquals($filePath, $xPath, $attribute, $value)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertEquals($value, $nodeList->item(0)->getAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     * @param mixed $value
     */
    public function assertZipXmlAttributeStartsWith($filePath, $xPath, $attribute, $value)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertStringStartsWith($value, $nodeList->item(0)->getAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     * @param mixed $value
     */
    public function assertZipXmlAttributeEndsWith($filePath, $xPath, $attribute, $value)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertStringEndsWith($value, $nodeList->item(0)->getAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     * @param mixed $value
     */
    public function assertZipXmlAttributeContains($filePath, $xPath, $attribute, $value)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertContains($value, $nodeList->item(0)->getAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     */
    public function assertZipXmlAttributeExists($filePath, $xPath, $attribute)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertTrue($nodeList->item(0)->hasAttribute($attribute));
    }

    /**
     * @param string $filePath
     * @param string $xPath
     * @param string $attribute
     */
    public function assertZipXmlAttributeNotExists($filePath, $xPath, $attribute)
    {
        $this->writePresentationFile($this->oPresentation, $this->writerName);
        $nodeList = $this->getXmlNodeList($filePath, $xPath);
        self::assertFalse($nodeList->item(0)->hasAttribute($attribute));
    }

    public function assertIsSchemaECMA376Valid()
    {
        // validate all XML files
        $path = realpath($this->workDirectory . '/ppt');
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($path));

        foreach ($iterator as $file) {
            /** @var \SplFileInfo $file */
            if ($file->getExtension() !== 'xml') {
                continue;
            }

            $fileName = str_replace('\\', '/', substr($file->getRealPath(), strlen($path) + 1));
            $dom = $this->getXmlDom('ppt/' . $fileName);
            $xmlSource = $dom->saveXML();

            $dom->loadXML($xmlSource);
            $dom->schemaValidate(__DIR__ . '/../../../resources/schema/ecma-376/pml.xsd');

            $error = libxml_get_last_error();
            if ($error instanceof \LibXMLError) {
                $this->failXmlError($error, $fileName, $xmlSource);
            }
        }
        unset($iterator);
    }

    public function assertIsSchemaOOXMLValid()
    {
        // validate all XML files
        $path = realpath($this->workDirectory . '/ppt');
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($path));

        foreach ($iterator as $file) {
            /** @var \SplFileInfo $file */
            if ($file->getExtension() !== 'xml') {
                continue;
            }

            $fileName = str_replace('\\', '/', substr($file->getRealPath(), strlen($path) + 1));
            $dom = $this->getXmlDom('ppt/' . $fileName);
            $xmlSource = $dom->saveXML();
            // In the ISO/ECMA standard the namespace has changed from
            // http://schemas.openxmlformats.org/ to http://purl.oclc.org/ooxml/
            // We need to use the http://purl.oclc.org/ooxml/ namespace to validate
            // the xml against the current schema
            $xmlSource = str_replace(array(
                "http://schemas.openxmlformats.org/drawingml/2006/main",
                "http://schemas.openxmlformats.org/drawingml/2006/chart",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "http://schemas.openxmlformats.org/presentationml/2006/main",
            ), array(
                "http://purl.oclc.org/ooxml/drawingml/main",
                "http://purl.oclc.org/ooxml/drawingml/chart",
                "http://purl.oclc.org/ooxml/officeDocument/relationships",
                "http://purl.oclc.org/ooxml/presentationml/main",
            ), $xmlSource);

            $dom->loadXML($xmlSource);
            $dom->schemaValidate(__DIR__ . '/../../../resources/schema/ooxml/pml.xsd');

            $error = libxml_get_last_error();
            if ($error instanceof \LibXMLError) {
                $this->failXmlError($error, $fileName, $xmlSource);
            }
        }
        unset($iterator);
    }

    /**
     * @param string $version
     * @param boolean $triggerError
     * @return boolean
     */
    public function assertIsSchemaOpenDocumentValid($version = '1.0', $triggerError = true)
    {
        if (!array_key_exists($version, $this->arrayOpenDocumentRNG)) {
            self::fail('assertIsSchemaOpenDocumentValid > Use a valid version');
            return;
        }

        // validate all XML files
        $path = realpath($this->workDirectory);
        $iterator = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($path));

        $isValid = true;
        foreach ($iterator as $file) {
            /** @var \SplFileInfo $file */
            if ($file->getExtension() !== 'xml') {
                continue;
            }

            $fileName = str_replace('\\', '/', substr($file->getRealPath(), strlen($path) + 1));
            $dom = $this->getXmlDom($fileName);
            $xmlSource = $dom->saveXML();
            
            $dom->loadXML($xmlSource);
            $pathRNG = __DIR__ . '/../../../resources/schema/opendocument/'.$version.'/';
            if (isset($this->arrayOpenDocumentRNG[$version][$fileName])) {
                $pathRNG .= $this->arrayOpenDocumentRNG[$version][$fileName];
            } else {
                $pathRNG .= $this->arrayOpenDocumentRNG[$version]['*'];
            }
            $dom->relaxNGValidate($pathRNG);

            $error = libxml_get_last_error();
            if ($error instanceof \LibXMLError) {
                if ($triggerError) {
                    $this->failXmlError($error, $fileName, $xmlSource, array('version' => $version));
                }
                $isValid = false;
            }
        }
        unset($iterator);
        return $isValid;
    }

    public function assertIsSchemaOpenDocumentNotValid($version = '1.0')
    {
        $isValid = $this->assertIsSchemaOpenDocumentValid($version, false);
        if ($isValid) {
            self::fail('Failed : This document is currently valid (Schema version: '.$version.')');
        }
    }

    /**
     * @param \LibXMLError $error
     * @param string $fileName
     * @param string $source
     * @param array $params
     */
    protected function failXmlError(\LibXMLError $error, $fileName, $source, array $params = array())
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
        $errorLine = (int)$error->line;
        $contents = explode("\n", $source);
        $lines = array();
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
            $paramStr .= "\n" . ' - Parameters :'."\n";
            foreach ($params as $key => $val) {
                $paramStr .= '   - '.$key.' : '.$val."\n";
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
