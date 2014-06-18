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

namespace PhpOffice\PhpPowerpoint\Shape;

use PhpOffice\PhpPowerpoint\Shape;
use PhpOffice\PhpPowerpoint\IComparable;
use PhpOffice\PhpPowerpoint\Style\Border;

/**
 * PHPPowerPoint_Shape_Line
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Line extends Shape implements IComparable
{
    /**
     * Create a new PHPPowerPoint_Shape_Line instance
     *
     * @param int $fromX
     * @param int $fromY
     * @param int $toX
     * @param int $toY
     */
    public function __construct($fromX, $fromY, $toX, $toY)
    {
        parent::__construct();
        $this->getBorder()->setLineStyle(Border::LINE_SINGLE);

        $this->setOffsetX($fromX);
        $this->setOffsetY($fromY);
        $this->setWidth($toX - $fromX);
        $this->setHeight($toY - $fromY);

    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->getBorder()->getLineStyle() . parent::getHashCode() . __CLASS__);
    }
}
