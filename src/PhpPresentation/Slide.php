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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation;

use PhpOffice\PhpPresentation\Slide\AbstractSlide;
use PhpOffice\PhpPresentation\Slide\Note;
use PhpOffice\PhpPresentation\Slide\SlideLayout;

/**
 * Slide class.
 */
class Slide extends AbstractSlide implements ComparableInterface, ShapeContainerInterface
{
    /**
     * The slide is shown in presentation.
     *
     * @var bool
     */
    protected $isVisible = true;

    /**
     * Slide layout.
     *
     * @var SlideLayout|null
     */
    private $slideLayout;

    /**
     * Slide master id.
     *
     * @var int
     */
    private $slideMasterId = 1;

    /**
     * @var Note
     */
    private $slideNote;

    /**
     * @var \PhpOffice\PhpPresentation\Slide\Animation[]
     */
    protected $animations = [];

    /**
     * Name of the title.
     *
     * @var string|null
     */
    protected $name;

    /**
     * Create a new slide.
     *
     * @param PhpPresentation $pParent
     */
    public function __construct(PhpPresentation $pParent = null)
    {
        // Set parent
        $this->parent = $pParent;
        // Shape collection
        $this->shapeCollection = new \ArrayObject();
        // Set identifier
        $this->identifier = md5(rand(0, 9999) . time());
        // Set Slide Layout
        if ($this->parent instanceof PhpPresentation) {
            $arrayMasterSlides = $this->parent->getAllMasterSlides();
            $oMasterSlide = reset($arrayMasterSlides);
            $arraySlideLayouts = $oMasterSlide->getAllSlideLayouts();
            $oSlideLayout = reset($arraySlideLayouts);
            $this->setSlideLayout($oSlideLayout);
        }
        // Set note
        $this->setNote(new Note());
    }

    /**
     * Get slide layout.
     */
    public function getSlideLayout(): ?SlideLayout
    {
        return $this->slideLayout;
    }

    /**
     * Set slide layout.
     */
    public function setSlideLayout(SlideLayout $layout): self
    {
        $this->slideLayout = $layout;

        return $this;
    }

    /**
     * Get slide master id.
     *
     * @return int
     */
    public function getSlideMasterId()
    {
        return $this->slideMasterId;
    }

    /**
     * Set slide master id.
     *
     * @param int $masterId
     *
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function setSlideMasterId($masterId = 1)
    {
        $this->slideMasterId = $masterId;

        return $this;
    }

    /**
     * Copy slide (!= clone!).
     *
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function copy()
    {
        $copied = clone $this;

        return $copied;
    }

    public function getNote(): Note
    {
        return $this->slideNote;
    }

    public function setNote(Note $note = null): self
    {
        $this->slideNote = (is_null($note) ? new Note() : $note);
        $this->slideNote->setParent($this);

        return $this;
    }

    /**
     * Get the name of the slide.
     *
     * @return string
     */
    public function getName(): ?string
    {
        return $this->name;
    }

    /**
     * Set the name of the slide.
     */
    public function setName(?string $name = null): self
    {
        $this->name = $name;

        return $this;
    }

    /**
     * @return bool
     */
    public function isVisible()
    {
        return $this->isVisible;
    }

    /**
     * @param bool $value
     *
     * @return Slide
     */
    public function setIsVisible($value = true)
    {
        $this->isVisible = (bool) $value;

        return $this;
    }

    /**
     * Add an animation to the slide.
     *
     * @param \PhpOffice\PhpPresentation\Slide\Animation $animation
     *
     * @return Slide
     */
    public function addAnimation($animation)
    {
        $this->animations[] = $animation;

        return $this;
    }

    /**
     * Get collection of animations.
     *
     * @return \PhpOffice\PhpPresentation\Slide\Animation[]
     */
    public function getAnimations()
    {
        return $this->animations;
    }

    /**
     * Set collection of animations.
     *
     * @param \PhpOffice\PhpPresentation\Slide\Animation[] $array
     *
     * @return Slide
     */
    public function setAnimations(array $array = [])
    {
        $this->animations = $array;

        return $this;
    }
}
