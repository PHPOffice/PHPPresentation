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

namespace PhpOffice\PhpPresentation\Slide;

use IteratorIterator;
use PhpOffice\PhpPresentation\PhpPresentation;
use ReturnTypeWillChange;

// @phpstan-ignore-next-line
class Iterator extends IteratorIterator
{
    /**
     * Presentation to iterate.
     *
     * @var PhpPresentation
     */
    private $subject;

    /**
     * Current iterator position.
     *
     * @var int
     */
    private $position = 0;

    /**
     * Create a new slide iterator.
     */
    public function __construct(PhpPresentation $subject)
    {
        $this->subject = $subject;
    }

    /**
     * Rewind iterator.
     */
    #[ReturnTypeWillChange]
    public function rewind(): void
    {
        $this->position = 0;
    }

    /**
     * Current \PhpOffice\PhpPresentation\Slide.
     *
     * @return \PhpOffice\PhpPresentation\Slide
     */
    #[ReturnTypeWillChange]
    public function current()
    {
        return $this->subject->getSlide($this->position);
    }

    /**
     * Current key.
     *
     * @return int
     */
    #[ReturnTypeWillChange]
    public function key()
    {
        return $this->position;
    }

    /**
     * Next value.
     */
    #[ReturnTypeWillChange]
    public function next(): void
    {
        ++$this->position;
    }

    /**
     * More \PhpOffice\PhpPresentation\Slide instances available?
     *
     * @return bool
     */
    #[ReturnTypeWillChange]
    public function valid()
    {
        return $this->position < $this->subject->getSlideCount();
    }
}
