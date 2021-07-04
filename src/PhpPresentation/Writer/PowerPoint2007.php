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

namespace PhpOffice\PhpPresentation\Writer;

use DirectoryIterator;
use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007.
 */
class PowerPoint2007 extends AbstractWriter implements WriterInterface
{
    /**
     * Use disk caching where possible?
     *
     * @var bool
     */
    protected $useDiskCaching = false;

    /**
     * Disk caching directory.
     *
     * @var string
     */
    protected $diskCachingDir;

    /**
     * Create a new PowerPoint2007 file.
     *
     * @param PhpPresentation $pPhpPresentation
     *
     * @throws \Exception
     */
    public function __construct(PhpPresentation $pPhpPresentation = null)
    {
        // Assign PhpPresentation
        $this->setPhpPresentation($pPhpPresentation);

        // Set up disk caching location
        $this->diskCachingDir = './';

        // Set HashTable variables
        $this->oDrawingHashTable = new HashTable();

        $this->setZipAdapter(new ZipArchiveAdapter());
    }

    /**
     * Save PhpPresentation to file.
     *
     * @throws \Exception
     */
    public function save(string $pFilename): void
    {
        if (empty($pFilename)) {
            throw new \Exception('Filename is empty');
        }
        $oPresentation = $this->getPhpPresentation();

        // If $pFilename is php://output or php://stdout, make it a temporary file...
        $originalFilename = $pFilename;
        if ('php://output' == strtolower($pFilename) || 'php://stdout' == strtolower($pFilename)) {
            $pFilename = @tempnam('./', 'phppttmp');
            if ('' == $pFilename) {
                $pFilename = $originalFilename;
            }
        }

        // Create drawing dictionary
        $this->getDrawingHashTable()->addFromSource($this->allDrawings());

        $oZip = $this->getZipAdapter();
        $oZip->open($pFilename);

        $oDir = new DirectoryIterator(dirname(__FILE__) . DIRECTORY_SEPARATOR . 'PowerPoint2007');
        $arrayFiles = [];
        foreach ($oDir as $oFile) {
            if (!$oFile->isFile()) {
                continue;
            }

            $class = __NAMESPACE__ . '\\PowerPoint2007\\' . $oFile->getBasename('.php');
            $class = new \ReflectionClass($class);

            if ($class->isAbstract() || !$class->isSubclassOf('PhpOffice\PhpPresentation\Writer\PowerPoint2007\AbstractDecoratorWriter')) {
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
            $oZip = $oService->render();
            unset($oService);
        }

        // Close file
        $oZip->close();

        // If a temporary file was used, copy it to the correct file stream
        if ($originalFilename != $pFilename) {
            if (false === copy($pFilename, $originalFilename)) {
                throw new \Exception("Could not copy temporary zip file $pFilename to $originalFilename.");
            }
            if (false === @unlink($pFilename)) {
                throw new \Exception('The file ' . $pFilename . ' could not be removed.');
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
     * @param string $pDirectory Disk caching directory
     *
     * @throws \Exception
     *
     * @return \PhpOffice\PhpPresentation\Writer\PowerPoint2007
     */
    public function setUseDiskCaching($pValue = false, $pDirectory = null)
    {
        $this->useDiskCaching = $pValue;

        if (!is_null($pDirectory)) {
            if (!is_dir($pDirectory)) {
                throw new \Exception("Directory does not exist: $pDirectory");
            }
            $this->diskCachingDir = $pDirectory;
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
        return $this->diskCachingDir;
    }
}
