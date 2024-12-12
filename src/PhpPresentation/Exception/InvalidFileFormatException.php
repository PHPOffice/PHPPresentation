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

namespace PhpOffice\PhpPresentation\Exception;

class InvalidFileFormatException extends PhpPresentationException
{
    public function __construct(string $path, string $class, string $error = '')
    {
        if ($class) {
            $class = 'class ' . $class;
        }
        if ($error) {
            $error = '(' . $error . ')';
        }

        parent::__construct(sprintf(
            'The file %s is not in the format supported by %s%s%s',
            $path,
            $class,
            !empty($error) ? ' ' : '',
            $error
        ));
    }
}
