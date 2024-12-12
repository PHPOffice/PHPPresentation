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

class UnauthorizedMimetypeException extends PhpPresentationException
{
    /**
     * @param array<string> $authorizedMimetypes
     */
    public function __construct(string $expectedMimetype, array $authorizedMimetypes)
    {
        parent::__construct(sprintf(
            'The mime type %s is not found in autorized values (%s)',
            $expectedMimetype,
            implode(', ', $authorizedMimetypes)
        ));
    }
}
