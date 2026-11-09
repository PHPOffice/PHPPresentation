<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PhpPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PhpPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL v3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Reader\PowerPoint2007;

use PhpOffice\PhpPresentation\Reader\PowerPoint2007;

/**
 * Mock class for the reader, with its answers about the paths of the parts of a package public.
 */
class PackagePathReader extends PowerPoint2007
{
    /**
     * Public wrapper for a protected method.
     */
    public function folderOf(string $part): string
    {
        return parent::folderOf($part);
    }

    /**
     * Public wrapper for a protected method.
     */
    public function relsOf(string $part): string
    {
        return parent::relsOf($part);
    }

    /**
     * Public wrapper for a protected method.
     */
    public function resolve(string $relsPart, string $target): string
    {
        return parent::resolve($relsPart, $target);
    }
}
