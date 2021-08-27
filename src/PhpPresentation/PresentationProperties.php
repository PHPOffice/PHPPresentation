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

class PresentationProperties
{
    public const VIEW_HANDOUT = 'handoutView';
    public const VIEW_NOTES = 'notesView';
    public const VIEW_NOTES_MASTER = 'notesMasterView';
    public const VIEW_OUTLINE = 'outlineView';
    public const VIEW_SLIDE = 'sldView';
    public const VIEW_SLIDE_MASTER = 'sldMasterView';
    public const VIEW_SLIDE_SORTER = 'sldSorterView';
    public const VIEW_SLIDE_THUMBNAIL = 'sldThumbnailView';

    /**
     * @var array<int, string>
     */
    protected $arrayView = [
        self::VIEW_HANDOUT,
        self::VIEW_NOTES,
        self::VIEW_NOTES_MASTER,
        self::VIEW_OUTLINE,
        self::VIEW_SLIDE,
        self::VIEW_SLIDE_MASTER,
        self::VIEW_SLIDE_SORTER,
        self::VIEW_SLIDE_THUMBNAIL,
    ];

    public const SLIDESHOW_TYPE_PRESENT = 'present';
    public const SLIDESHOW_TYPE_BROWSE = 'browse';
    public const SLIDESHOW_TYPE_KIOSK = 'kiosk';

    /**
     * @var array<int, string>
     */
    protected $arraySlideshowTypes = [
        self::SLIDESHOW_TYPE_PRESENT,
        self::SLIDESHOW_TYPE_BROWSE,
        self::SLIDESHOW_TYPE_KIOSK,
    ];

    /**
     * @var bool
     */
    protected $isLoopUntilEsc = false;

    /**
     * Mark as final.
     *
     * @var bool
     */
    protected $markAsFinal = false;

    /**
     * @var string|null
     */
    protected $thumbnail;

    /**
     * Zoom.
     *
     * @var float
     */
    protected $zoom = 1.0;

    /**
     * @var string
     */
    protected $lastView = self::VIEW_SLIDE;

    /**
     * @var string
     */
    protected $slideshowType = self::SLIDESHOW_TYPE_PRESENT;

    /**
     * @var bool
     */
    protected $isCommentVisible = false;

    public function isLoopContinuouslyUntilEsc(): bool
    {
        return $this->isLoopUntilEsc;
    }

    public function setLoopContinuouslyUntilEsc(bool $value = false): self
    {
        $this->isLoopUntilEsc = $value;

        return $this;
    }

    /**
     * Return the thumbnail file path.
     *
     * @return string|null
     */
    public function getThumbnailPath(): ?string
    {
        return $this->thumbnail;
    }

    /**
     * Define the path for the thumbnail file / preview picture.
     *
     * @param string $path
     *
     * @return self
     */
    public function setThumbnailPath(string $path = ''): self
    {
        if (file_exists($path)) {
            $this->thumbnail = $path;
        }

        return $this;
    }

    /**
     * Mark a document as final.
     */
    public function markAsFinal(bool $state = true): self
    {
        $this->markAsFinal = $state;

        return $this;
    }

    /**
     * Return if this document is marked as final.
     *
     * @return bool
     */
    public function isMarkedAsFinal(): bool
    {
        return $this->markAsFinal;
    }

    /**
     * Set the zoom of the document (in percentage).
     */
    public function setZoom(float $zoom = 1.0): self
    {
        $this->zoom = $zoom;

        return $this;
    }

    /**
     * Return the zoom (in percentage).
     */
    public function getZoom(): float
    {
        return $this->zoom;
    }

    /**
     * @param string $value
     *
     * @return self
     */
    public function setLastView(string $value = self::VIEW_SLIDE): self
    {
        if (in_array($value, $this->arrayView)) {
            $this->lastView = $value;
        }

        return $this;
    }

    /**
     * @return string
     */
    public function getLastView(): string
    {
        return $this->lastView;
    }

    public function setCommentVisible(bool $value = false): self
    {
        $this->isCommentVisible = $value;

        return $this;
    }

    public function isCommentVisible(): bool
    {
        return $this->isCommentVisible;
    }

    /**
     * @return string
     */
    public function getSlideshowType(): string
    {
        return $this->slideshowType;
    }

    /**
     * @param string $value
     *
     * @return self
     */
    public function setSlideshowType(string $value = self::SLIDESHOW_TYPE_PRESENT): self
    {
        if (in_array($value, $this->arraySlideshowTypes)) {
            $this->slideshowType = $value;
        }

        return $this;
    }
}
