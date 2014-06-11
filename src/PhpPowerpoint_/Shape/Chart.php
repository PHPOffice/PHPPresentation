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

use PhpOffice\PhpPowerpoint\IComparable;
use PhpOffice\PhpPowerpoint\Shape\BaseDrawing;
use PhpOffice\PhpPowerpoint\Shape\Chart\Title;
use PhpOffice\PhpPowerpoint\Shape\Chart\Legend;
use PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea;
use PhpOffice\PhpPowerpoint\Shape\Chart\View3D;

/**
 * PHPPowerPoint_Shape_Chart
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Shape
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class Chart extends BaseDrawing implements IComparable
{
    /**
     * Title
     *
     * @var PHPPowerPoint_Shape_Chart_Title
     */
    private $_title;

    /**
     * Legend
     *
     * @var PHPPowerPoint_Shape_Chart_Legend
     */
    private $_legend;

    /**
     * Plot area
     *
     * @var PHPPowerPoint_Shape_Chart_PlotArea
     */
    private $_plotArea;

    /**
     * View 3D
     *
     * @var PHPPowerPoint_Shape_Chart_View3D
     */
    private $_view3D;

    /**
     * Include spreadsheet for editing data? Requires PHPExcel in the same folder as PHPPowerPoint
     *
     * @var bool
     */
    private $_includeSpreadsheet = false;

    /**
     * Create a new PHPPowerPoint_Slide_MemoryDrawing
     */
    public function __construct()
    {
        // Initialize
        $this->_title    = new Title();
        $this->_legend   = new Legend();
        $this->_plotArea = new PlotArea();
        $this->_view3D   = new View3D();

        // Initialize parent
        parent::__construct();
    }

    /**
     * Get Title
     *
     * @return PHPPowerPoint_Shape_Chart_Title
     */
    public function getTitle()
    {
        return $this->_title;
    }

    /**
     * Get Legend
     *
     * @return PHPPowerPoint_Shape_Chart_Legend
     */
    public function getLegend()
    {
        return $this->_legend;
    }

    /**
     * Get PlotArea
     *
     * @return PHPPowerPoint_Shape_Chart_PlotArea
     */
    public function getPlotArea()
    {
        return $this->_plotArea;
    }

    /**
     * Get View3D
     *
     * @return PHPPowerPoint_Shape_Chart_View3D
     */
    public function getView3D()
    {
        return $this->_view3D;
    }

    /**
     * Include spreadsheet for editing data? Requires PHPExcel in the same folder as PHPPowerPoint
     *
     * @return boolean
     */
    public function getIncludeSpreadsheet()
    {
        return $this->_includeSpreadsheet;
    }

    /**
     * Include spreadsheet for editing data? Requires PHPExcel in the same folder as PHPPowerPoint
     *
     * @param  boolean                   $value
     * @return PHPPowerPoint_Shape_Chart
     */
    public function setIncludeSpreadsheet($value = false)
    {
        $this->_includeSpreadsheet = $value;

        return $this;
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename()
    {
        return 'chart' . $this->getImageIndex() . '.xml';
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5(parent::getHashCode() . $this->_title->getHashCode() . $this->_legend->getHashCode() . $this->_plotArea->getHashCode() . $this->_view3D->getHashCode() . ($this->_includeSpreadsheet ? 1 : 0) . __CLASS__);
    }

    /**
     * Hash index
     *
     * @var string
     */
    private $_hashIndex;

    /**
     * Get hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return string Hash index
     */
    public function getHashIndex()
    {
        return $this->_hashIndex;
    }

    /**
     * Set hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param string $value Hash index
     */
    public function setHashIndex($value)
    {
        $this->_hashIndex = $value;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
}
