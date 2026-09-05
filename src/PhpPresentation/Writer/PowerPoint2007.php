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

use DirectoryIterator;
use DK\OpenXml\Encryption\AgileEncryptionOptions;
use DK\OpenXml\Encryption\EncryptedOfficeFile;
use DK\OpenXml\OpenXmlPackage;
use PhpOffice\PhpPresentation\Exception\DirectoryNotFoundException;
use PhpOffice\PhpPresentation\Exception\FileCopyException;
use PhpOffice\PhpPresentation\Exception\FileRemoveException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\AbstractDecoratorWriter;
use ReflectionClass;

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
     */
    public function __construct(?PhpPresentation $pPhpPresentation = null)
    {
        // Assign PhpPresentation
        $this->setPhpPresentation($pPhpPresentation ?? new PhpPresentation());

        // Set up disk caching location
        $this->diskCachingDir = './';

        // Set HashTable variables
        $this->oDrawingHashTable = new HashTable();
    }

    /**
     * The password to lock the presentation with.
     *
     * @var null|string
     */
    protected $encryptionPassword;

    /**
     * The password-hash iterations to lock it with.
     *
     * @var int
     */
    protected $encryptionSpinCount = 100000;

    /**
     * Save PhpPresentation to file.
     */
    public function save(string $pFilename): void
    {
        if (empty($pFilename)) {
            throw new InvalidParameterException('pFilename', '');
        }
        $oPresentation = $this->getPhpPresentation();

        // If $pFilename is php://output or php://stdout, make it a temporary file...
        $originalFilename = $pFilename;
        if ('php://output' == strtolower($pFilename) || 'php://stdout' == strtolower($pFilename)) {
            $pFilename = @tempnam($this->diskCachingDir, 'phppttmp');
            if ('' == $pFilename) {
                $pFilename = $originalFilename;
            }
        }

        // Create drawing dictionary
        $this->getDrawingHashTable()->addFromSource($this->allDrawings());

        $this->numberMasterSlides();

        $package = OpenXmlPackage::create();
        // an extension that carries one type throughout is said once for the package rather than
        // once for every part written under it, which is what Office writes
        foreach (AbstractDecoratorWriter::MEDIA_CONTENT_TYPES as $extension => $contentType) {
            $package->setDefaultContentType($extension, $contentType);
        }

        $oDir = new DirectoryIterator(__DIR__ . DIRECTORY_SEPARATOR . 'PowerPoint2007');
        $arrayFiles = [];
        foreach ($oDir as $oFile) {
            if (!$oFile->isFile()) {
                continue;
            }

            $class = __NAMESPACE__ . '\\PowerPoint2007\\' . $oFile->getBasename('.php');
            $class = new ReflectionClass($class);

            if ($class->isAbstract() || !$class->isSubclassOf(AbstractDecoratorWriter::class)) {
                continue;
            }
            $arrayFiles[$oFile->getBasename('.php')] = $class;
        }

        ksort($arrayFiles);

        foreach ($arrayFiles as $o) {
            $oService = $o->newInstance();
            $oService->setPackage($package);
            $oService->setPresentation($oPresentation);
            $oService->setDrawingHashTable($this->getDrawingHashTable());
            $package = $oService->render();
            unset($oService);
        }

        $package->saveAs($pFilename);

        $this->encrypt($pFilename);

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
     * Number the masters and the layouts before anything is written.
     *
     * A part names another by the number in its filename -- `slideLayout3.xml` -- so every
     * decorator has to agree on which layout is the third, and a layout carries an identifier of
     * its own that the presentation part lists. Nothing derives these from the model, so they are
     * settled once, here, rather than by whichever decorator happens to run first.
     */
    private function numberMasterSlides(): void
    {
        $layoutNr = 0;
        $layoutId = time() + 689016272; // requires minimum value of 2 147 483 648

        foreach ($this->getPhpPresentation()->getAllMasterSlides() as $idx => $oSlideMaster) {
            $oSlideMaster->setRelsIndex((string) ($idx + 1));
            foreach ($oSlideMaster->getAllSlideLayouts() as $oSlideLayout) {
                $oSlideLayout->layoutNr = ++$layoutNr;
                $oSlideLayout->setRelsIndex((string) $oSlideLayout->layoutNr);
                $oSlideLayout->layoutId = ++$layoutId;
            }
        }
    }

    /**
     * Lock the presentation with a password.
     */
    public function setEncryptionPassword(string $encryptionPassword): self
    {
        $this->encryptionPassword = $encryptionPassword;

        return $this;
    }

    /**
     * Say which Agile profile to lock the presentation with.
     *
     * Only AES-256 with SHA-512 is written, which is what Office writes and the strongest of the
     * profiles ECMA-376 allows; the weaker ones are read but not produced. The signature takes
     * them anyway so that a caller who names a profile is told it cannot be written rather than
     * given a different one silently, and so that it reads the same as PhpSpreadsheet's.
     */
    public function setEncryptionProfile(int $keyBits, string $hashAlgorithm, int $spinCount = 100000): self
    {
        if (256 !== $keyBits || 0 !== strcasecmp($hashAlgorithm, 'SHA512')) {
            throw new InvalidParameterException(
                'keyBits, hashAlgorithm',
                $keyBits . ', ' . $hashAlgorithm,
                'only AES-256 with SHA512 is written'
            );
        }
        $this->encryptionSpinCount = $spinCount;

        return $this;
    }

    /**
     * A protected presentation is the package wrapped in a compound file, which is written from
     * the package rather than in place, so it is written beside it and moved over it.
     */
    protected function encrypt(string $pFilename): void
    {
        if (null === $this->encryptionPassword) {
            return;
        }

        $locked = $pFilename . '.locked';
        EncryptedOfficeFile::encrypt(
            $pFilename,
            $locked,
            $this->encryptionPassword,
            new AgileEncryptionOptions($this->encryptionSpinCount)
        );

        if (false === @rename($locked, $pFilename)) {
            @unlink($locked);

            throw new FileCopyException($locked, $pFilename);
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
     * @return PowerPoint2007
     */
    public function setUseDiskCaching(bool $useDiskCaching = false, ?string $directory = null)
    {
        $this->useDiskCaching = $useDiskCaching;

        if (null !== $directory) {
            if (!is_dir($directory)) {
                throw new DirectoryNotFoundException($directory);
            }
            $this->diskCachingDir = $directory;
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
