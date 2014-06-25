<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPowerpoint\Slide;

use PhpOffice\PhpPowerpoint\PhpPowerpoint;

/**
 * \PhpOffice\PhpPowerpoint\Slide\Iterator
 *
 * Used to iterate slides in PHPPowerPoint
 */
class Iterator extends \IteratorIterator
{
    /**
     * Presentation to iterate
     *
     * @var \PhpOffice\PhpPowerpoint\PhpPowerpoint
     */
    private $subject;

    /**
     * Current iterator position
     *
     * @var int
     */
    private $position = 0;

    /**
     * Create a new slide iterator
     *
     * @param PHPPowerPoint $subject
     */
    public function __construct(PhpPowerpoint $subject = null)
    {
        // Set subject
        $this->subject = $subject;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset($this->subject);
    }

    /**
     * Rewind iterator
     */
    public function rewind()
    {
        $this->position = 0;
    }

    /**
     * Current \PhpOffice\PhpPowerpoint\Slide
     *
     * @return \PhpOffice\PhpPowerpoint\Slide
     */
    public function current()
    {
        return $this->subject->getSlide($this->position);
    }

    /**
     * Current key
     *
     * @return int
     */
    public function key()
    {
        return $this->position;
    }

    /**
     * Next value
     */
    public function next()
    {
        ++$this->position;
    }

    /**
     * More \PhpOffice\PhpPowerpoint\Slide instances available?
     *
     * @return boolean
     */
    public function valid()
    {
        return $this->position < $this->subject->getSlideCount();
    }
}
