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

class InvalidParameterException extends PhpPresentationException
{
    public function __construct(string $parameter, string $value, ?string $error = null)
    {
        $message = sprintf(
            'The parameter %s can\'t have the value "%s"',
            $parameter,
            $value
        );
        if ($error) {
            $message = sprintf('%s (Validation: %s)', $message, $error);
        }

        parent::__construct($message);
    }
}
