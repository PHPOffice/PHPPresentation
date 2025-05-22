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

use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\Exception\WriterPDFAdapterNotDefinedException;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;

class PDF extends HTML implements WriterInterface
{
    /**
     * Create a new \PhpOffice\PhpPresentation\Writer\ODPresentation.
     */
    public function __construct(?PhpPresentation $pPhpPresentation = null)
    {
        // Assign PhpPresentation
        $this->setPhpPresentation($pPhpPresentation ?? new PhpPresentation());

        // Set HashTable variables
        $this->oDrawingHashTable = new HashTable();
    }

    /**
     * Save PhpPresentation to file.
     */
    public function save(string $pFilename): void
    {
        if (empty($pFilename)) {
            throw new InvalidParameterException('pFilename', '');
        }

        if (empty($this->getPDFAdapter())) {
            throw new WriterPDFAdapterNotDefinedException();
        }

        $this->getPDFAdapter()->setPhpPresentation($this->getPhpPresentation());
        $this->getPDFAdapter()->save($pFilename);
    }
}
