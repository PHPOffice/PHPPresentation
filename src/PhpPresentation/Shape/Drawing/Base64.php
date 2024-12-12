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

namespace PhpOffice\PhpPresentation\Shape\Drawing;

use PhpOffice\PhpPresentation\Exception\UnauthorizedMimetypeException;

class Base64 extends AbstractDrawingAdapter
{
    /**
     * @var string
     */
    protected $data;

    /**
     * Unique name.
     *
     * @var string
     */
    protected $uniqueName;

    /**
     * @var array<string, string>
     */
    protected $arrayMimeExtension = [
        'image/jpeg' => 'jpg',
        'image/png' => 'png',
        'image/gif' => 'gif',
        'image/svg+xml' => 'svg',
    ];

    /**
     * @var string
     */
    protected $path;

    /**
     * Base64 constructor.
     */
    public function __construct()
    {
        parent::__construct();
        $this->uniqueName = md5(mt_rand(0, 9999) . time() . mt_rand(0, 9999));
        $this->data = '';
    }

    public function getData(): string
    {
        return $this->data;
    }

    public function setData(string $data): self
    {
        $this->data = $data;

        return $this;
    }

    public function getContents(): string
    {
        [, $imageContents] = explode(';', $this->getData());
        [, $imageContents] = explode(',', $imageContents);

        return base64_decode($imageContents);
    }

    public function getExtension(): string
    {
        [$data] = explode(';', $this->getData());
        [, $mime] = explode(':', $data);

        if (!array_key_exists($mime, $this->arrayMimeExtension)) {
            throw new UnauthorizedMimetypeException($mime, $this->arrayMimeExtension);
        }

        return $this->arrayMimeExtension[$mime];
    }

    public function getIndexedFilename(): string
    {
        return $this->uniqueName . $this->getImageIndex() . '.' . $this->getExtension();
    }

    public function getMimeType(): string
    {
        [$data] = explode(';', $this->getData());
        [, $mime] = explode(':', $data);

        if (!empty($mime)) {
            return $mime;
        }

        $sImage = $this->getContents();
        if (!function_exists('getimagesizefromstring')) {
            $uri = 'data://application/octet-stream;base64,' . base64_encode($sImage);
            $image = getimagesize($uri);
        } else {
            $image = getimagesizefromstring($sImage);
        }

        return image_type_to_mime_type($image[2]);
    }

    /**
     * Get Path.
     */
    public function getPath(): string
    {
        return $this->path;
    }

    public function setPath(string $path): self
    {
        $this->path = $path;

        return $this;
    }
}
