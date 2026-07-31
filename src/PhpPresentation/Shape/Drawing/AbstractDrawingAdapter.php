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

use PhpOffice\PhpPresentation\Shape\AbstractGraphic;

abstract class AbstractDrawingAdapter extends AbstractGraphic
{
    abstract public function getContents(): string;

    abstract public function getExtension(): string;

    abstract public function getIndexedFilename(): string;

    abstract public function getMimeType(): string;

    abstract public function getPath(): string;

    /**
     * @param string $path File path
     * @return self
     */
    abstract public function setPath(string $path);

    /**
     * Set whether this is a temporary file that should be cleaned up
     *
     * @param bool $isTemporary
     * @return self
     */
    abstract public function setIsTemporaryFile(bool $isTemporary);

    /**
     * Load content into this object using a temporary file
     *
     * @param string $content Binary content
     * @param string $fileName Optional fileName for reference
     * @param string $prefix Prefix for the temporary file
     * @return self
     */
    public function loadFromContent(string $content, string $fileName = '', string $prefix = 'PhpPresentation'): self
    {
        $tmpFile = tempnam(sys_get_temp_dir(), $prefix);
        file_put_contents($tmpFile, $content);

        // Set path and mark as temporary for automatic cleanup
        $this->setPath($tmpFile);
        $this->setIsTemporaryFile(true);

        // Set filename if provided
        if (!empty($fileName)) {
            $this->setName($fileName);
        }

        return $this;
    }
}
