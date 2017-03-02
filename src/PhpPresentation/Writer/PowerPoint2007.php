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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Writer;

use DirectoryIterator;
use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\AbstractLayoutPack;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\PackDefault;

/**
 * \PhpOffice\PhpPresentation\Writer\PowerPoint2007
 */
class PowerPoint2007 extends AbstractWriter implements WriterInterface
{
    /**
     * Use disk caching where possible?
     *
     * @var boolean
     */
    protected $useDiskCaching = false;

    /**
     * Disk caching directory
     *
     * @var string
     */
    protected $diskCachingDir;

    /**
     * Layout pack to use
     * @deprecated 0.7
     * @var \PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\AbstractLayoutPack
     */
    protected $layoutPack;

    /**
     * Create a new PowerPoint2007 file
     *
     * @param PhpPresentation $pPhpPresentation
     */
    public function __construct(PhpPresentation $pPhpPresentation = null)
    {
        // Assign PhpPresentation
        $this->setPhpPresentation($pPhpPresentation);

        // Set up disk caching location
        $this->diskCachingDir = './';

        // Set layout pack
        $this->layoutPack = new PackDefault();

        // Set HashTable variables
        $this->oDrawingHashTable = new HashTable();

        $this->setZipAdapter(new ZipArchiveAdapter());
    }

    /**
     * Save PhpPresentation to file
     *
     * @param  string    $pFilename
     * @throws \Exception
     */
    public function save($pFilename)
    {
        if (empty($pFilename)) {
            throw new \Exception("Filename is empty");
        }
        $oPresentation = $this->getPhpPresentation();

        // If $pFilename is php://output or php://stdout, make it a temporary file...
        $originalFilename = $pFilename;
        if (strtolower($pFilename) == 'php://output' || strtolower($pFilename) == 'php://stdout') {
            $pFilename = @tempnam('./', 'phppttmp');
            if ($pFilename == '') {
                $pFilename = $originalFilename;
            }
        }

        // Create drawing dictionary
        $this->getDrawingHashTable()->addFromSource($this->allDrawings());

        $oZip = $this->getZipAdapter();
        $oZip->open($pFilename);

        $oDir = new DirectoryIterator(dirname(__FILE__).DIRECTORY_SEPARATOR.'PowerPoint2007');
        $arrayFiles = array();
        foreach ($oDir as $oFile) {
            if (!$oFile->isFile()) {
                continue;
            }

            $class = __NAMESPACE__ . '\\PowerPoint2007\\' . $oFile->getBasename('.php');
            $o = new \ReflectionClass($class);

            if ($o->isAbstract() || !$o->isSubclassOf('PhpOffice\PhpPresentation\Writer\PowerPoint2007\AbstractDecoratorWriter')) {
                continue;
            }
            $arrayFiles[$oFile->getBasename('.php')] = $o;
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
            if (copy($pFilename, $originalFilename) === false) {
                throw new \Exception("Could not copy temporary zip file $pFilename to $originalFilename.");
            }
            if (@unlink($pFilename) === false) {
                throw new \Exception('The file '.$pFilename.' could not be removed.');
            }
        }
    }

    /**
     * Get use disk caching where possible?
     *
     * @return boolean
     */
    public function hasDiskCaching()
    {
        return $this->useDiskCaching;
    }

    /**
     * Set use disk caching where possible?
     *
     * @param  boolean $pValue
     * @param  string $pDirectory Disk caching directory
     * @throws \Exception
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
     * Get disk caching directory
     *
     * @return string
     */
    public function getDiskCachingDirectory()
    {
        return $this->diskCachingDir;
    }

    /**
     * Get layout pack to use
     *
     * @deprecated 0.7
     * @return \PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\AbstractLayoutPack
     */
    public function getLayoutPack()
    {
        return $this->layoutPack;
    }

    /**
     * Set layout pack to use
     *
     * @deprecated 0.7
     * @param \PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\AbstractLayoutPack $pValue
     * @return \PhpOffice\PhpPresentation\Writer\PowerPoint2007
     */
    public function setLayoutPack(AbstractLayoutPack $pValue = null)
    {
        $this->layoutPack = $pValue;

        return $this;
    }
}
