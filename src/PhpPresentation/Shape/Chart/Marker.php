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

namespace PhpOffice\PhpPresentation\Shape\Chart;

use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Border;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Axis
 */
class Marker
{
    const SYMBOL_CIRCLE = 'circle';
    const SYMBOL_DASH = 'dash';
    const SYMBOL_DIAMOND = 'diamond';
    const SYMBOL_DOT = 'dot';
    const SYMBOL_NONE = 'none';
    const SYMBOL_PLUS = 'plus';
    const SYMBOL_SQUARE = 'square';
    const SYMBOL_STAR = 'star';
    const SYMBOL_TRIANGLE = 'triangle';
    const SYMBOL_X = 'x';

    public static $arraySymbol = array(
        self::SYMBOL_CIRCLE,
        self::SYMBOL_DASH,
        self::SYMBOL_DIAMOND,
        self::SYMBOL_DOT,
        self::SYMBOL_NONE,
        self::SYMBOL_PLUS,
        self::SYMBOL_SQUARE,
        self::SYMBOL_STAR,
        self::SYMBOL_TRIANGLE,
        self::SYMBOL_X
    );

    /**
     * @var string
     */
    protected $symbol = self::SYMBOL_NONE;

    /**
     * @var int
     */
    protected $size = 5;
	
	 /**
     * @var PhpOffice\PhpPresentation\Style\Fill
     */
    protected $fill;
	
	 /**
     * @var PhpOffice\PhpPresentation\Style\Border
     */
    protected $border;
	
	public function __construct(){
		$this->fill = new Fill();
		$this->border = new Border();
		
	}

    /**
     * @return string
     */
    public function getSymbol()
    {
        return $this->symbol;
    }

    /**
     * @param string $symbol
     * @return $this
     */
    public function setSymbol($symbol = self::SYMBOL_NONE)
    {
        $this->symbol = $symbol;
        return $this;
    }

    /**
     * @return int
     */
    public function getSize()
    {
        return $this->size;
    }

    /**
     * @param int $size
     * @return $this
     */
    public function setSize($size = 5)
    {
        $this->size = $size;
        return $this;
    }
	
	/**
     * @return PhpOffice\PhpPresentation\Style\Fill
     */
    public function getFill()
    {
        return $this->fill;
    }

    /**
     * @param PhpOffice\PhpPresentation\Style\Fill
     * @return $this
     */
    public function setFill(Fill $fill)
    {
        $this->fill = $fill;
        return $this;
    }
	
	/**
     * @return PhpOffice\PhpPresentation\Style\Border
     */
    public function getBorder()
    {
        return $this->border;
    }

    /**
     * @param PhpOffice\PhpPresentation\Style\Border
     * @return $this
     */
    public function setBorder(Border $border)
    {
        $this->border = $border;
        return $this;
    }
}
