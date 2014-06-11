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
 * PHPPowerPoint_Slide_Iterator
 *
 * Used to iterate slides in PHPPowerPoint
 */
class Iterator extends \IteratorIterator
{
    /**
     * Presentation to iterate
     *
     * @var PHPPowerPoint
     */
    private $_subject;

    /**
     * Current iterator position
     *
     * @var int
     */
    private $_position = 0;

    /**
     * Create a new slide iterator
     *
     * @param PHPPowerPoint $subject
     */
    public function __construct(PhpPowerpoint $subject = null)
    {
        // Set subject
        $this->_subject = $subject;
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        unset($this->_subject);
    }

    /**
     * Rewind iterator
     */
    public function rewind()
    {
        $this->_position = 0;
    }

    /**
     * Current PHPPowerPoint_Slide
     *
     * @return PHPPowerPoint_Slide
     */
    public function current()
    {
        return $this->_subject->getSlide($this->_position);
    }

    /**
     * Current key
     *
     * @return int
     */
    public function key()
    {
        return $this->_position;
    }

    /**
     * Next value
     */
    public function next()
    {
        ++$this->_position;
    }

    /**
     * More PHPPowerPoint_Slide instances available?
     *
     * @return boolean
     */
    public function valid()
    {
        return $this->_position < $this->_subject->getSlideCount();
    }
}
