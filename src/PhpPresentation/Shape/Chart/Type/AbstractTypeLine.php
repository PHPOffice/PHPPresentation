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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Shape\Chart\Type;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Type\Line
 */
abstract class AbstractTypeLine extends AbstractType implements ComparableInterface
{
	/**
	* Is Line Smooth?
	* @var boolean 
	*/
   protected $isSmooth = false;
   
	/**
	* Is Line Smooth?
	*
	* @return boolean
	*/
   public function getIsSmooth(){
	   return $this->isSmooth;
   }
   
    /**
    * Set Line Smoothness
    *
    * @param  boolean $value
    * @return $this
    */
    public function setIsSmooth($value = true)
    {
        $this->isSmooth = $value;
        return $this;
    }
   
   
}
