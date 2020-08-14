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

/**
 * \PhpOffice\PhpPresentation\PresentationProperties
 */
class PresentationProperties
{
    const VIEW_HANDOUT = 'handoutView';
    const VIEW_NOTES = 'notesView';
    const VIEW_NOTES_MASTER = 'notesMasterView';
    const VIEW_OUTLINE = 'outlineView';
    const VIEW_SLIDE = 'sldView';
    const VIEW_SLIDE_MASTER = 'sldMasterView';
    const VIEW_SLIDE_SORTER = 'sldSorterView';
    const VIEW_SLIDE_THUMBNAIL = 'sldThumbnailView';
    const THUMBNAIL_FILE = 'file'; // Thumbnail path is out of PPT
    const THUMBNAIL_ZIP = 'zip'; // Thumbnail path point to an image store into file loaded

    protected $arrayView = array(
        self::VIEW_HANDOUT,
        self::VIEW_NOTES,
        self::VIEW_NOTES_MASTER,
        self::VIEW_OUTLINE,
        self::VIEW_SLIDE,
        self::VIEW_SLIDE_MASTER,
        self::VIEW_SLIDE_SORTER,
        self::VIEW_SLIDE_THUMBNAIL,
    );

    /*
     * @var boolean
     */
    protected $isLoopUntilEsc = false;

    /**
     * Mark as final
     * @var bool
     */
    protected $markAsFinal = false;

    /*
     * @var string Define the thumbnail content (if content into zip file)
     */
    protected $thumbnail = null;

    /*
     * @var string Define the thumbnail place
     */
    protected $thumbnailPath = '';

    /*
     * @var string Define if thumbnail is out of PPT or previouly store into PPT
     */
    protected $thumbnailType = self::THUMBNAIL_FILE;

    /**
     * Zoom
     * @var float
     */
    protected $zoom = 1;

    /*
     * @var string
     */
    protected $lastView = self::VIEW_SLIDE;

    /*
     * @var boolean
     */
    protected $isCommentVisible = false;
    
    /**
     * @return bool
     */
    public function isLoopContinuouslyUntilEsc()
    {
        return $this->isLoopUntilEsc;
    }
    
    /**
     * @param bool $value
     * @return \PhpOffice\PhpPresentation\PresentationProperties
     */
    public function setLoopContinuouslyUntilEsc($value = false)
    {
        if (is_bool($value)) {
            $this->isLoopUntilEsc = $value;
        }
        return $this;
    }
    
    /**
     * Return the thumbnail file path
     * @return string
     */
    public function getThumbnailPath()
    {
        return $this->thumbnailPath;
    }
    
    /**
     * Return the content of thumbnail
     *
     * @return binary Content of image
     */
    public function getThumbnail()
    {
      // Return content of local file
      if ($this->getThumbnailType() == self::THUMBNAIL_FILE) {
        if (file_exists($this->getThumbnailPath()))
          return file_get_contents($this->getThumbnailPath());
      }
      // Return content of image stored into zip file
      if ($this->getThumbnailType() == self::THUMBNAIL_ZIP) {
        return $this->thumbnail;
      }
      // Return null if no thumbnail
      return null;
    }
    
    /**
     * Define the path for the thumbnail file / preview picture
     * @param string $path
     * @return \PhpOffice\PhpPresentation\PresentationProperties
     */
    public function setThumbnailPath($path = '', $type = self::THUMBNAIL_FILE, $content = null)
    {
        if (file_exists($path) && ($type == self::THUMBNAIL_FILE)) {
            $this->thumbnailPath = $path;
            $this->thumbnailType = $type;
        }
        if (($path != '') && ($type == self::THUMBNAIL_ZIP)) {
            $this->thumbnailPath = $path;
            $this->thumbnailType = $type;
            $this->thumbnail = $content;
        }
        return $this;
    }

    /**
     * Return the thumbnail type
     * @return string
     */
    public function getThumbnailType()
    {
        return $this->thumbnailType;
    }
    
    /**
     * Mark a document as final
     * @param bool $state
     * @return PresentationProperties
     */
    public function markAsFinal($state = true)
    {
        if (is_bool($state)) {
            $this->markAsFinal = $state;
        }
        return $this;
    }

    /**
     * Return if this document is marked as final
     * @return bool
     */
    public function isMarkedAsFinal()
    {
        return $this->markAsFinal;
    }

    /**
     * Set the zoom of the document (in percentage)
     * @param float $zoom
     * @return PresentationProperties
     */
    public function setZoom($zoom = 1.0)
    {
        if (is_numeric($zoom)) {
            $this->zoom = (float)$zoom;
        }
        return $this;
    }

    /**
     * Return the zoom (in percentage)
     * @return float
     */
    public function getZoom()
    {
        return $this->zoom;
    }

    /**
     * @param string $value
     * @return $this
     */
    public function setLastView($value = self::VIEW_SLIDE)
    {
        if (in_array($value, $this->arrayView)) {
            $this->lastView = $value;
        }
        return $this;
    }

    /**
     * @return string
     */
    public function getLastView()
    {
        return $this->lastView;
    }

    /**
     * @param bool $value
     * @return $this
     */
    public function setCommentVisible($value = false)
    {
        if (is_bool($value)) {
            $this->isCommentVisible = $value;
        }
        return $this;
    }

    /**
     * @return string
     */
    public function isCommentVisible()
    {
        return $this->isCommentVisible;
    }
}
