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

namespace PhpOffice\PhpPresentation\Writer;

use DirectoryIterator;
use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;
use PhpOffice\PhpPresentation\Exception\DirectoryNotFoundException;
use PhpOffice\PhpPresentation\Exception\FileCopyException;
use PhpOffice\PhpPresentation\Exception\FileRemoveException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;

/**
 * ODPresentation writer.
 */
class ODPresentation extends AbstractWriter implements WriterInterface
{
    /**
     * @var \PhpOffice\PhpPresentation\Shape\Chart[]
     */
    public $chartArray = [];

    /**
     * Use disk caching where possible?
     *
     * @var bool
     */
    private $useDiskCaching = false;

    /**
     * Disk caching directory.
     *
     * @var string
     */
    private $diskCachingDirectory;

    /**
     * Create a new \PhpOffice\PhpPresentation\Writer\ODPresentation.
     *
     * @param PhpPresentation $pPhpPresentation
     */
    public function __construct(PhpPresentation $pPhpPresentation = null)
    {
        // Assign PhpPresentation
        $this->setPhpPresentation($pPhpPresentation ?? new PhpPresentation());

        // Set up disk caching location
        $this->diskCachingDirectory = './';

        // Set HashTable variables
        $this->oDrawingHashTable = new HashTable();

        $this->setZipAdapter(new ZipArchiveAdapter());
    }

    /**
     * Save PhpPresentation to file.
     *
     * @throws FileCopyException
     * @throws FileRemoveException
     * @throws InvalidParameterException
     */
    public function save(string $pFilename): void
    {
        if (empty($pFilename)) {
            throw new InvalidParameterException('pFilename', '');
        }
        // If $pFilename is php://output or php://stdout, make it a temporary file...
        $originalFilename = $pFilename;
        if ('php://output' == strtolower($pFilename) || 'php://stdout' == strtolower($pFilename)) {
            $pFilename = @tempnam('./', 'phppttmp');
            if ('' == $pFilename) {
                $pFilename = $originalFilename;
            }
        }

        // Initialize HashTable
        $this->getDrawingHashTable()->addFromSource($this->allDrawings());

        // Initialize Zip
        $oZip = $this->getZipAdapter();
        $oZip->open($pFilename);

        // Variables
        $oPresentation = $this->getPhpPresentation();
        $arrayChart = [];

        $arrayFiles = [];
        $oDir = new DirectoryIterator(dirname(__FILE__) . DIRECTORY_SEPARATOR . 'ODPresentation');
        foreach ($oDir as $oFile) {
            if (!$oFile->isFile()) {
                continue;
            }

            $class = __NAMESPACE__ . '\\ODPresentation\\' . $oFile->getBasename('.php');
            $class = new \ReflectionClass($class);

            if ($class->isAbstract() || !$class->isSubclassOf('PhpOffice\PhpPresentation\Writer\ODPresentation\AbstractDecoratorWriter')) {
                continue;
            }
            $arrayFiles[$oFile->getBasename('.php')] = $class;
        }

        ksort($arrayFiles);

        foreach ($arrayFiles as $o) {
            $oService = $o->newInstance();
            $oService->setZip($oZip);
            $oService->setPresentation($oPresentation);
            $oService->setDrawingHashTable($this->getDrawingHashTable());
            $oService->setArrayChart($arrayChart);
            $oZip = $oService->render();
            $arrayChart = $oService->getArrayChart();
            unset($oService);
        }

        // Close file
        $oZip->close();

        // If a temporary file was used, copy it to the correct file stream
        if ($originalFilename != $pFilename) {
            if (false === copy($pFilename, $originalFilename)) {
                throw new FileCopyException($pFilename, $originalFilename);
            }
            if (false === @unlink($pFilename)) {
                throw new FileRemoveException($pFilename);
            }
        }
    }

    /**
     * Get use disk caching where possible?
     *
     * @return bool
     */
    public function hasDiskCaching()
    {
        return $this->useDiskCaching;
    }

    /**
     * Set use disk caching where possible?
     *
     * @param bool $pValue
     * @param string $directory Disk caching directory
     *
     * @throws DirectoryNotFoundException
     *
     * @return \PhpOffice\PhpPresentation\Writer\ODPresentation
     */
    public function setUseDiskCaching(bool $pValue = false, string $directory = null)
    {
        $this->useDiskCaching = $pValue;

        if (!is_null($directory)) {
            if (!is_dir($directory)) {
                throw new DirectoryNotFoundException($directory);
            }
            $this->diskCachingDirectory = $directory;
        }

        return $this;
    }

    /**
     * Get disk caching directory.
     *
     * @return string
     */
    public function getDiskCachingDirectory()
    {
        return $this->diskCachingDirectory;
    }
}
