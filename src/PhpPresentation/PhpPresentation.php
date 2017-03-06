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

namespace PhpOffice\PhpPresentation;

use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Iterator;
use PhpOffice\PhpPresentation\Slide\SlideMaster;

/**
 * PhpPresentation
 */
class PhpPresentation
{
    /**
     * Document properties
     *
     * @var \PhpOffice\PhpPresentation\DocumentProperties
     */
    protected $documentProperties;

    /**
     * Presentation properties
     *
     * @var \PhpOffice\PhpPresentation\PresentationProperties
     */
    protected $presentationProps;

    /**
     * Document layout
     *
     * @var \PhpOffice\PhpPresentation\DocumentLayout
     */
    protected $layout;

    /**
     * Collection of Slide objects
     *
     * @var \PhpOffice\PhpPresentation\Slide[]
     */
    protected $slideCollection = array();

    /**
     * Active slide index
     *
     * @var int
     */
    protected $activeSlideIndex = 0;

    /**
     * Collection of Master Slides
     * @var \ArrayObject|\PhpOffice\PhpPresentation\Slide\SlideMaster[]
     */
    protected $slideMasters;

    /**
     * Create a new PhpPresentation with one Slide
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
     * Get properties
     *
     * @return \PhpOffice\PhpPresentation\DocumentProperties
     * @deprecated for getDocumentProperties
     */
    public function getProperties()
    {
        return $this->getDocumentProperties();
    }

    /**
     * Set properties
     *
     * @param  \PhpOffice\PhpPresentation\DocumentProperties $value
     * @deprecated for setDocumentProperties
     * @return PhpPresentation
     */
    public function setProperties(DocumentProperties $value)
    {
        return $this->setDocumentProperties($value);
    }

    /**
     * Get properties
     *
     * @return \PhpOffice\PhpPresentation\DocumentProperties
     */
    public function getDocumentProperties()
    {
        return $this->documentProperties;
    }

    /**
     * Set properties
     *
     * @param  \PhpOffice\PhpPresentation\DocumentProperties $value
     * @return PhpPresentation
     */
    public function setDocumentProperties(DocumentProperties $value)
    {
        $this->documentProperties = $value;

        return $this;
    }

    /**
     * Get presentation properties
     *
     * @return \PhpOffice\PhpPresentation\PresentationProperties
     */
    public function getPresentationProperties()
    {
        return $this->presentationProps;
    }

    /**
     * Set presentation properties
     *
     * @param  \PhpOffice\PhpPresentation\PresentationProperties $value
     * @return PhpPresentation
     */
    public function setPresentationProperties(PresentationProperties $value)
    {
        $this->presentationProps = $value;
        return $this;
    }

    /**
     * Get layout
     *
     * @return \PhpOffice\PhpPresentation\DocumentLayout
     */
    public function getLayout()
    {
        return $this->layout;
    }

    /**
     * Set layout
     *
     * @param  \PhpOffice\PhpPresentation\DocumentLayout $value
     * @return PhpPresentation
     */
    public function setLayout(DocumentLayout $value)
    {
        $this->layout = $value;

        return $this;
    }

    /**
     * Get active slide
     *
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function getActiveSlide()
    {
        return $this->slideCollection[$this->activeSlideIndex];
    }

    /**
     * Create slide and add it to this presentation
     *
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function createSlide()
    {
        $newSlide = new Slide($this);
        $this->addSlide($newSlide);
        return $newSlide;
    }

    /**
     * Add slide
     *
     * @param  \PhpOffice\PhpPresentation\Slide $slide
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function addSlide(Slide $slide = null)
    {
        $this->slideCollection[] = $slide;

        return $slide;
    }

    /**
     * Remove slide by index
     *
     * @param  int $index Slide index
     * @throws \Exception
     * @return PhpPresentation
     */
    public function removeSlideByIndex($index = 0)
    {
        if ($index > count($this->slideCollection) - 1) {
            throw new \Exception("Slide index is out of bounds.");
        } else {
            array_splice($this->slideCollection, $index, 1);
        }

        return $this;
    }

    /**
     * Get slide by index
     *
     * @param  int $index Slide index
     * @return \PhpOffice\PhpPresentation\Slide
     * @throws \Exception
     */
    public function getSlide($index = 0)
    {
        if ($index > count($this->slideCollection) - 1) {
            throw new \Exception("Slide index is out of bounds.");
        } else {
            return $this->slideCollection[$index];
        }
    }

    /**
     * Get all slides
     *
     * @return \PhpOffice\PhpPresentation\Slide[]
     */
    public function getAllSlides()
    {
        return $this->slideCollection;
    }

    /**
     * Get index for slide
     *
     * @param  \PhpOffice\PhpPresentation\Slide\AbstractSlide $slide
     * @return int
     * @throws \Exception
     */
    public function getIndex(Slide\AbstractSlide $slide)
    {
        $index = null;
        foreach ($this->slideCollection as $key => $value) {
            if ($value->getHashCode() == $slide->getHashCode()) {
                $index = $key;
                break;
            }
        }
        return $index;
    }

    /**
     * Get slide count
     *
     * @return int
     */
    public function getSlideCount()
    {
        return count($this->slideCollection);
    }

    /**
     * Get active slide index
     *
     * @return int Active slide index
     */
    public function getActiveSlideIndex()
    {
        return $this->activeSlideIndex;
    }

    /**
     * Set active slide index
     *
     * @param  int $index Active slide index
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function setActiveSlideIndex($index = 0)
    {
        if ($index > count($this->slideCollection) - 1) {
            throw new \Exception("Active slide index is out of bounds.");
        } else {
            $this->activeSlideIndex = $index;
        }

        return $this->getActiveSlide();
    }

    /**
     * Add external slide
     *
     * @param  \PhpOffice\PhpPresentation\Slide $slide External slide to add
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Slide
     */
    public function addExternalSlide(Slide $slide)
    {
        $slide->rebindParent($this);

        $this->addMasterSlide($slide->getSlideLayout()->getSlideMaster());

        return $this->addSlide($slide);
    }

    /**
     * Get slide iterator
     *
     * @return \PhpOffice\PhpPresentation\Slide\Iterator
     */
    public function getSlideIterator()
    {
        return new Iterator($this);
    }

    /**
     * Create a masterslide and add it to this presentation
     *
     * @return \PhpOffice\PhpPresentation\Slide\SlideMaster
     */
    public function createMasterSlide()
    {
        $newMasterSlide = new SlideMaster($this);
        $this->addMasterSlide($newMasterSlide);
        return $newMasterSlide;
    }

    /**
     * Add masterslide
     *
     * @param  \PhpOffice\PhpPresentation\Slide\SlideMaster $slide
     * @return \PhpOffice\PhpPresentation\Slide\SlideMaster
     * @throws \Exception
     */
    public function addMasterSlide(SlideMaster $slide = null)
    {
        $this->slideMasters[] = $slide;

        return $slide;
    }

    /**
     * Copy presentation (!= clone!)
     *
     * @return PhpPresentation
     */
    public function copy()
    {
        $copied = clone $this;

        $slideCount = count($this->slideCollection);
        for ($i = 0; $i < $slideCount; ++$i) {
            $this->slideCollection[$i] = $this->slideCollection[$i]->copy();
            $this->slideCollection[$i]->rebindParent($this);
        }

        return $copied;
    }

    /**
     * Mark a document as final
     * @param bool $state
     * @return PhpPresentation
     * @deprecated for getPresentationProperties()->markAsFinal()
     */
    public function markAsFinal($state = true)
    {
        return $this->getPresentationProperties()->markAsFinal($state);
    }

    /**
     * Return if this document is marked as final
     * @return bool
     * @deprecated for getPresentationProperties()->isMarkedAsFinal()
     */
    public function isMarkedAsFinal()
    {
        return $this->getPresentationProperties()->isMarkedAsFinal();
    }

    /**
     * Set the zoom of the document (in percentage)
     * @param float $zoom
     * @return PhpPresentation
     * @deprecated for getPresentationProperties()->setZoom()
     */
    public function setZoom($zoom = 1)
    {
        return $this->getPresentationProperties()->setZoom($zoom);
    }

    /**
     * Return the zoom (in percentage)
     * @return float
     * @deprecated for getPresentationProperties()->getZoom()
     */
    public function getZoom()
    {
        return $this->getPresentationProperties()->getZoom();
    }

    /**
     * @return \ArrayObject|Slide\SlideMaster[]
     */
    public function getAllMasterSlides()
    {
        return $this->slideMasters;
    }

    /**
     * @param \ArrayObject|Slide\SlideMaster[] $slideMasters
     * @return $this
     */
    public function setAllMasterSlides($slideMasters = array())
    {
        if ($slideMasters instanceof \ArrayObject || is_array($slideMasters)) {
            $this->slideMasters = $slideMasters;
        }
        return $this;
    }
}
