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

namespace PhpOffice\PhpPresentation\Shared;

use PhpOffice\WMF\Reader\Detector;
use PhpOffice\WMF\Reader\EMF\GD as EMFReader;
use PhpOffice\WMF\Reader\WMF\GD as WMFReader;
use Throwable;

/**
 * Windows metafiles (WMF, EMF & EMF+), based on the optional library phpoffice/wmf.
 *
 * `getimagesize()` knows none of them, so this is what measures them and says what they are.
 * PowerPoint and LibreOffice store them as they are; a writer whose target can't show them
 * (HTML, PDF, Keynote) converts them to PNG.
 *
 * An EMF+ file is an EMF file holding EMF+ records, and is stored as an EMF file.
 */
class Metafile
{
    public const MIME_WMF = 'image/x-wmf';
    public const MIME_EMF = 'image/x-emf';

    /**
     * Returns whether the metafiles are supported : phpoffice/wmf is installed and GD is loaded.
     */
    public static function isSupported(): bool
    {
        return class_exists(Detector::class) && extension_loaded('gd');
    }

    /**
     * Returns the mime type of a metafile (`image/x-wmf` or `image/x-emf`), or null if the
     * contents are no metafile or the metafiles are not supported.
     */
    public static function getMimeType(string $contents): ?string
    {
        if (!self::isSupported()) {
            return null; // @codeCoverageIgnore
        }

        switch (Detector::detect($contents)) {
            case Detector::TYPE_WMF:
                return self::MIME_WMF;
            case Detector::TYPE_EMF:
            case Detector::TYPE_EMFPLUS:
                return self::MIME_EMF;
            default:
                return null;
        }
    }

    /**
     * Returns whether a mime type is the one of a metafile, with or without its `x-`.
     */
    public static function isMimeType(string $mimeType): bool
    {
        return null !== self::getExtension($mimeType);
    }

    /**
     * Returns the extension of a metafile mime type (`wmf` or `emf`), or null.
     */
    public static function getExtension(string $mimeType): ?string
    {
        switch (strtolower($mimeType)) {
            case self::MIME_WMF:
            case 'image/wmf':
                return 'wmf';
            case self::MIME_EMF:
            case 'image/emf':
                return 'emf';
            default:
                return null;
        }
    }

    /**
     * Returns the size of a metafile and its mime type, as `getimagesize()` does : [width, height, mimeType].
     *
     * Returns null if the contents are no metafile, can't be read, or the metafiles are not supported.
     *
     * @return null|array{0: int, 1: int, 2: string}
     */
    public static function getImageSize(string $contents): ?array
    {
        $mimeType = self::getMimeType($contents);
        if (null === $mimeType) {
            return null;
        }
        $reader = self::load($contents, $mimeType);
        if (null === $reader) {
            return null;
        }
        $image = $reader->getResource();

        return [imagesx($image), imagesy($image), $mimeType];
    }

    /**
     * Converts a metafile to PNG.
     *
     * Returns null if the contents are no metafile, can't be read, or the metafiles are not supported.
     */
    public static function convertToPng(string $contents): ?string
    {
        $mimeType = self::getMimeType($contents);
        if (null === $mimeType) {
            return null;
        }
        $reader = self::load($contents, $mimeType);
        if (null === $reader) {
            return null;
        }

        ob_start();
        $isWritten = imagepng($reader->getResource());
        $png = (string) ob_get_clean();

        return $isWritten && '' !== $png ? $png : null;
    }

    /**
     * @return null|EMFReader|WMFReader
     */
    private static function load(string $contents, string $mimeType)
    {
        $reader = self::MIME_WMF === $mimeType ? new WMFReader() : new EMFReader();

        try {
            if (!$reader->loadFromString($contents)) {
                return null;
            }
        } catch (Throwable $e) {
            return null;
        }

        return $reader;
    }
}
