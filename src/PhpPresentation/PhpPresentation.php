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
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation;

use ArrayObject;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\Slide\Iterator;
use PhpOffice\PhpPresentation\Slide\SlideMaster;

/**
 * PhpPresentation.
 */
class PhpPresentation
{
    /**
     * Document properties.
     *
     * @var DocumentProperties
     */
    protected $documentProperties;

    /**
     * Presentation properties.
     *
     * @var PresentationProperties
     */
    protected $presentationProps;

    /**
     * Document layout.
     *
     * @var DocumentLayout
     */
    protected $layout;

    /**
     * Collection of Slide objects.
     *
     * @var array<int, Slide>
     */
    protected $slideCollection = [];

    /**
     * Active slide index.
     *
     * @var int
     */
    protected $activeSlideIndex = 0;

    /**
     * Collection of Master Slides.
     *
     * @var array<int, SlideMaster>|ArrayObject<int, SlideMaster>
     */
    protected $slideMasters;

    /**
     * Create a new PhpPresentation with one Slide.
     */
    public function __construct()
    {
        // Set empty Master & SlideLayout
        $this->createMasterSlide()->createSlideLayout();

        // Initialise slide collection and add one slide
        $this->createSlide();
        $this->setActiveSlideIndex();

        // Set initial document properties & layout
        $this->setDocumentProperties(new DocumentProperties());
        $this->setPresentationProperties(new PresentationProperties());
        $this->setLayout(new DocumentLayout());
    }

    /**
     * Get properties.
     */
    public function getDocumentProperties(): DocumentProperties
    {
        return $this->documentProperties;
    }

    /**
     * Set properties.
     */
    public function setDocumentProperties(DocumentProperties $value): self
    {
        $this->documentProperties = $value;

        return $this;
    }

    /**
     * Get presentation properties.
     */
    public function getPresentationProperties(): PresentationProperties
    {
        return $this->presentationProps;
    }

    /**
     * Set presentation properties.
     */
    public function setPresentationProperties(PresentationProperties $value): self
    {
        $this->presentationProps = $value;

        return $this;
    }

    /**
     * Get layout.
     */
    public function getLayout(): DocumentLayout
    {
        return $this->layout;
    }

    /**
     * Set layout.
     */
    public function setLayout(DocumentLayout $value): self
    {
        $this->layout = $value;

        return $this;
    }

    /**
     * Get active slide.
     */
    public function getActiveSlide(): Slide
    {
        return $this->slideCollection[$this->activeSlideIndex];
    }

    /**
     * Create slide and add it to this presentation.
     */
    public function createSlide(): Slide
    {
        $newSlide = new Slide($this);
        $this->addSlide($newSlide);

        return $newSlide;
    }

    /**
     * Add slide.
     */
    public function addSlide(Slide $slide, int $index = -1): Slide
    {
        if ($index > -1) {
            array_splice($this->slideCollection, $index, 0, [$slide]);
        } else {
            $this->slideCollection[] = $slide;
        }

        return $slide;
    }

    /**
     * Remove slide by index.
     *
     * @param int $index Slide index
     */
    public function removeSlideByIndex(int $index = 0): self
    {
        if ($index > count($this->slideCollection) - 1) {
            throw new OutOfBoundsException(0, count($this->slideCollection) - 1, $index);
        }
        array_splice($this->slideCollection, $index, 1);

        return $this;
    }

    /**
     * Get slide by index.
     *
     * @param int $index Slide index
     */
    public function getSlide(int $index = 0): Slide
    {
        if ($index > count($this->slideCollection) - 1) {
            throw new OutOfBoundsException(0, count($this->slideCollection) - 1, $index);
        }

        return $this->slideCollection[$index];
    }

    /**
     * Get all slides.
     *
     * @return array<int, Slide>
     */
    public function getAllSlides(): array
    {
        return $this->slideCollection;
    }

    /**
     * Get index for slide.
     */
    public function getIndex(Slide\AbstractSlide $slide): ?int
    {
        if (empty($this->slideCollection)) {
            return null;
        }
        foreach ($this->slideCollection as $key => $value) {
            if ($value->getHashCode() == $slide->getHashCode()) {
                return $key;
            }
        }

        return null;
    }

    /**
     * Get slide count.
     */
    public function getSlideCount(): int
    {
        return count($this->slideCollection);
    }

    /**
     * Get active slide index.
     *
     * @return int Active slide index
     */
    public function getActiveSlideIndex(): int
    {
        return $this->activeSlideIndex;
    }

    /**
     * Set active slide index.
     *
     * @param int $index Active slide index
     */
    public function setActiveSlideIndex(int $index = 0): Slide
    {
        if ($index > count($this->slideCollection) - 1) {
            throw new OutOfBoundsException(0, count($this->slideCollection) - 1, $index);
        }
        $this->activeSlideIndex = $index;

        return $this->getActiveSlide();
    }

    /**
     * Add external slide.
     *
     * @param Slide $slide External slide to add
     */
    public function addExternalSlide(Slide $slide): Slide
    {
        $slide->rebindParent($this);

        $this->addMasterSlide($slide->getSlideLayout()->getSlideMaster());

        return $this->addSlide($slide);
    }

    /**
     * Get slide iterator.
     */
    public function getSlideIterator(): Iterator
    {
        return new Iterator($this);
    }

    /**
     * Create a masterslide and add it to this presentation.
     */
    public function createMasterSlide(): SlideMaster
    {
        $newMasterSlide = new SlideMaster($this);
        $this->addMasterSlide($newMasterSlide);

        return $newMasterSlide;
    }

    /**
     * Add masterslide.
     */
    public function addMasterSlide(SlideMaster $slide): SlideMaster
    {
        $this->slideMasters[] = $slide;

        return $slide;
    }

    /**
     * Copy presentation (!= clone!).
     */
    public function copy(): self
    {
        $copied = clone $this;

        $slideCount = count($this->slideCollection);

        // Because the rebindParent() method on AbstractSlide removes the slide
        // from the parent (current $this which we're cloning) presentation, we
        // save the collection. This way, after the copying has finished, we can
        // return the slides to the original presentation.
        $oldSlideCollection = $this->slideCollection;
        $newSlideCollection = [];

        for ($i = 0; $i < $slideCount; ++$i) {
            $newSlideCollection[$i] = $oldSlideCollection[$i]->copy();
            $newSlideCollection[$i]->rebindParent($copied);
        }

        // Give the copied presentation a copied slide collection which the
        // copied slides have been rebind to the copied presentation.
        $copied->slideCollection = $newSlideCollection;

        // Return the original slides to the original presentation.
        $this->slideCollection = $oldSlideCollection;

        return $copied;
    }

    /**
     * @return array<int, SlideMaster>|ArrayObject<int, SlideMaster>
     */
    public function getAllMasterSlides()
    {
        return $this->slideMasters;
    }

    /**
     * @param array<int, SlideMaster>|ArrayObject<int, SlideMaster> $slideMasters
     */
    public function setAllMasterSlides($slideMasters = []): self
    {
        if ($slideMasters instanceof ArrayObject || is_array($slideMasters)) {
            $this->slideMasters = $slideMasters;
        }

        return $this;
    }
}
