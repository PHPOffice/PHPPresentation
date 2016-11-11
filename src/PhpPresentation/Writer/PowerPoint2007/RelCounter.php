<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;


class RelCounter implements \Iterator, \Countable
{
    private $i;

    public function __construct()
    {
        $this->i = 0;
    }

    /**
     * @inheritdoc
     */
    public function current()
    {
        return $this->i;
    }

    /**
     * @inheritdoc
     */
    public function next()
    {
        ++$this->i;
    }

    /**
     * @inheritdoc
     */
    public function key()
    {
        return $this->i;
    }

    /**
     * @inheritdoc
     */
    public function valid()
    {
        return true;
    }

    /**
     * @inheritdoc
     */
    public function rewind()
    {
        $this->i = 0;
    }

    /**
     * @inheritdoc
     */
    public function count()
    {
        return $this->i;
    }
}
