<?php
/**
 * Created by PhpStorm.
 * Author: Vincent Kool
 * User: vakool
 * Date: 14-03-16
 * Time: 14:56
 *
 * Description: Placeholders are created at the master and layout levels.
 * That is, they can be added and deleted at those levels but not at the slide level.
 * At the slide level they are only referenced or linked to the lower layout layer.
 * That is, a slide can reference a placeholder at the layout level but not at the master level.
 * Similarly a layout can reference a placeholder at the master slide level.
 *
 */
namespace PhpOffice\PhpPresentation\Shape;
class Placeholder
{
    /** Placeholder Type constants */
    const PH_TYPE_BODY = 'body';
    const PH_TYPE_CHART = 'chart';
    const PH_TYPE_SUBTITLE = 'subTitle';
    const PH_TYPE_TITLE = 'title';
    const PH_TYPE_FOOTER = 'ftr';
    const PH_TYPE_DATETIME = 'dt';
    const PH_TYPE_SLIDENUM = 'sldNum';
    /**
     * hasCustomPrompt
     * Indicates whether the placeholder should have a customer prompt.
     *
     * @var bool
     */
    protected $hasCustomPrompt;
    /**
     * idx
     * Specifies the index of the placeholder. This is used when applying templates or changing layouts to
     * match a placeholder on one template or master to another.
     *
     * @var int
     */
    protected $idx;
    /**
     * type
     * Specifies what content type the placeholder is to contains
     */
    protected $type;

    /**
     * Placeholder constructor.
     * @param $type
     */
    public function __construct($type)
    {
        $this->type = $type;
        return $this;
    }

    /**
     * @return mixed
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * @param mixed $type
     * @return string
     */
    public function setType($type)
    {
        $this->type = $type;
        return $this;
    }

    /**
     * @return int
     */
    public function getIdx()
    {
        return $this->idx;
    }

    /**
     * @param int $idx
     */
    public function setIdx($idx)
    {
        $this->idx = $idx;
    }
}