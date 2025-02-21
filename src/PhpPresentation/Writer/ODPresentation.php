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

namespace PhpOffice\PhpPresentation\Writer;

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
     */
    public function __construct(?PhpPresentation $pPhpPresentation = null)
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
     */
    public function save(string $pFilename): void
    {
        if (empty($pFilename)) {
            throw new InvalidParameterException('pFilename', '');
        }
        // If $pFilename is php://output or php://stdout, make it a temporary file...
        $originalFilename = $pFilename;
        if ('php://output' == strtolower($pFilename) || 'php://stdout' == strtolower($pFilename)) {
            $pFilename = @tempnam($this->diskCachingDirectory, 'phppttmp');
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

        foreach ([
            __CLASS__ . '\Mimetype',
            __CLASS__ . '\Content',
            __CLASS__ . '\Meta',
            __CLASS__ . '\MetaInfManifest',
            __CLASS__ . '\ObjectsChart',
            __CLASS__ . '\Pictures',
            __CLASS__ . '\Styles',
            __CLASS__ . '\ThumbnailsThumbnail',
        ] as $class) {
            $oService = new $class();
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
     * @param string $directory Disk caching directory
     *
     * @return ODPresentation
     */
    public function setUseDiskCaching(bool $pValue = false, ?string $directory = null)
    {
        $this->useDiskCaching = $pValue;

        if (null !== $directory) {
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
