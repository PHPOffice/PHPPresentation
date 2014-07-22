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

use PhpOffice\PhpPowerpoint\ComparableInterface;
use PhpOffice\PhpPowerpoint\Shape\Chart\Legend;
use PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea;
use PhpOffice\PhpPowerpoint\Shape\Chart\Title;
use PhpOffice\PhpPowerpoint\Shape\Chart\View3D;

/**
 * Chart element
 */
class Chart extends AbstractDrawing implements ComparableInterface
{
    /**
     * Title
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\Chart\Title
     */
    private $title;

    /**
     * Legend
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\Chart\Legend
     */
    private $legend;

    /**
     * Plot area
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea
     */
    private $plotArea;

    /**
     * View 3D
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\Chart\View3D
     */
    private $view3D;

    /**
     * Include spreadsheet for editing data? Requires PHPExcel in the same folder as PHPPowerPoint
     *
     * @var bool
     */
    private $includeSpreadsheet = false;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Slide\MemoryDrawing
     */
    public function __construct()
    {
        // Initialize
        $this->title    = new Title();
        $this->legend   = new Legend();
        $this->plotArea = new PlotArea();
        $this->view3D   = new View3D();

        // Initialize parent
        parent::__construct();
    }

    /**
     * Get Title
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Title
     */
    public function getTitle()
    {
        return $this->title;
    }

    /**
     * Get Legend
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\Legend
     */
    public function getLegend()
    {
        return $this->legend;
    }

    /**
     * Get PlotArea
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\PlotArea
     */
    public function getPlotArea()
    {
        return $this->plotArea;
    }

    /**
     * Get View3D
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart\View3D
     */
    public function getView3D()
    {
        return $this->view3D;
    }

    /**
     * Include spreadsheet for editing data? Requires PHPExcel in the same folder as PHPPowerPoint
     *
     * @return boolean
     */
    public function hasIncludedSpreadsheet()
    {
        return $this->includeSpreadsheet;
    }

    /**
     * Include spreadsheet for editing data? Requires PHPExcel in the same folder as PHPPowerPoint
     *
     * @param  boolean                   $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Chart
     */
    public function setIncludeSpreadsheet($value = false)
    {
        $this->includeSpreadsheet = $value;
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
        return md5(parent::getHashCode() . $this->title->getHashCode() . $this->legend->getHashCode() . $this->plotArea->getHashCode() . $this->view3D->getHashCode() . ($this->includeSpreadsheet ? 1 : 0) . __CLASS__);
    }
}
