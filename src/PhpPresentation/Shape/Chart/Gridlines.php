<?php
/**
 * Created by PhpStorm.
 * User: lefevre_f
 * Date: 02/03/2016
 * Time: 14:25
 */

namespace PhpPresentation\Shape\Chart;

class Gridlines
{
    /**
     * @var Outline
     */
    protected $outline;

    public function __construct()
    {
        $this->outline = new Outline();
    }

    /**
     * @return Outline
     */
    public function getOutline()
    {
        return $this->outline;
    }

    /**
     * @param Outline $outline
     * @return Gridlines
     */
    public function setOutline(Outline $outline)
    {
        $this->outline = $outline;
        return $this;
    }
}