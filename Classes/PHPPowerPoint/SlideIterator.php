<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_SlideIterator
 *
 * Used to iterate slides in PHPPowerPoint
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_SlideIterator extends IteratorIterator
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
	 * @param PHPPowerPoint 		$subject
	 */
	public function __construct(PHPPowerPoint $subject = null) {
		// Set subject
		$this->_subject = $subject;
	}

	/**
	 * Destructor
	 */
	public function __destruct() {
		unset($this->_subject);
	}

	/**
	 * Rewind iterator
	 */
    public function rewind() {
        $this->_position = 0;
    }

    /**
     * Current PHPPowerPoint_Slide
     *
     * @return PHPPowerPoint_Slide
     */
    public function current() {
    	return $this->_subject->getSlide($this->_position);
    }

    /**
     * Current key
     *
     * @return int
     */
    public function key() {
        return $this->_position;
    }

    /**
     * Next value
     */
    public function next() {
        ++$this->_position;
    }

    /**
     * More PHPPowerPoint_Slide instances available?
     *
     * @return boolean
     */
    public function valid() {
        return $this->_position < $this->_subject->getSlideCount();
    }
}
