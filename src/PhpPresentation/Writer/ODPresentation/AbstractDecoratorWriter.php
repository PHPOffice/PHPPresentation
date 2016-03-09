<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\PhpPresentation\Shape\Chart;

abstract class AbstractDecoratorWriter extends \PhpOffice\PhpPresentation\Writer\AbstractDecoratorWriter
{
    /**
     * @var Chart[]
     */
    protected $arrayChart;

    /**
     * @return \PhpOffice\PhpPresentation\Shape\Chart[]
     */
    public function getArrayChart()
    {
        return $this->arrayChart;
    }

    /**
     * @param \PhpOffice\PhpPresentation\Shape\Chart[] $arrayChart
     * @return AbstractDecoratorWriter
     */
    public function setArrayChart($arrayChart)
    {
        $this->arrayChart = $arrayChart;
        return $this;
    }
}
