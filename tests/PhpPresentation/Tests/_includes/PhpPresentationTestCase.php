<?php

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;

class PhpPresentationTestCase extends \PHPUnit_Framework_Assert
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
     * Executed before any method of the class, once
     */
    public function setUpBeforeClass()
    {
        $this->workDirectory = sys_get_temp_dir() . '/PhpPresentation_Unit_Test/';
    }

    /**
     * Executed before each method of the class
     */
    public function setUp()
    {
        $this->oPresentation = new PhpPresentation();

        $this->filePath = tempnam(sys_get_temp_dir(), 'PhpPresentation');
        if (!is_dir($this->workDirectory)) {
            mkdir($this->workDirectory);
        }
    }

    /**
     * Executed after each method of the class
     */
    public function tearDown()
    {
        $this->oPresentation = null;
        if (file_exists($this->filePath)) {
            unlink($this->filePath);
        }
        if (is_dir($this->workDirectory)) {
            $this->deleteDir($this->workDirectory);
        }
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

    /**
     * @param PhpPresentation $oPhpPresentation
     * @param string $writerName
     */
    private function writePresentationFile(PhpPresentation $oPhpPresentation, $writerName)
    {
        if (file_exists($this->filePath)) {
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

    /**
     * @param PhpPresentation $oPhpPresentation
     * @param string $writerName
     * @param string $filePath
     */
    public function assertZipFileExists(PhpPresentation $oPhpPresentation, $writerName, $filePath)
    {
        $this->writePresentationFile($oPhpPresentation, $writerName);
        self::assertThat(file_exists($this->workDirectory . $filePath), self::isTrue());
    }

    /**
     * @param PhpPresentation $oPhpPresentation
     * @param string $writerName
     * @param string $filePath
     */
    public function assertZipFileNotExists(PhpPresentation $oPhpPresentation, $writerName, $filePath)
    {
        $this->writePresentationFile($oPhpPresentation, $writerName);
        self::assertThat(file_exists($this->workDirectory . $filePath), self::isFalse());
    }
}
