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
     * @return self
     */
    abstract public function setPath(string $path);

    /**
     * The name a drawing is given in the package it is written to.
     *
     * The file name it was read from, and the index that tells two drawings with the
     * same file name apart: the writers write one media part per name, so two
     * pictures called `logo.png` would share one part, and one of them would come
     * out with the other's picture.
     */
    protected function nameFromPath(string $extension): string
    {
        $name = pathinfo($this->getPath(), PATHINFO_FILENAME);
        $name .= $this->getImageIndex();
        $name .= '.' . $extension;

        return str_replace(' ', '_', $name);
    }
}
