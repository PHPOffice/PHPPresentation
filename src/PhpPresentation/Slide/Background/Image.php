<?php

namespace PhpOffice\PhpPresentation\Slide\Background;

use PhpOffice\PhpPresentation\Slide\AbstractBackground;

class Image extends AbstractBackground
{
    /**
     * @var string
     */
    public $relationId;

    /**
     * Path
     *
     * @var string
     */
    protected $path;

    /**
     * @var int
     */
    protected $height;

    /**
     * @var int
     */
    protected $width;

    /**
     * Get Path
     *
     * @return string
     */
    public function getPath()
    {
        return $this->path;
    }

    /**
     * Set Path
     *
     * @param  string                      $pValue      File path
     * @param  boolean                     $pVerifyFile Verify file
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Shape\Drawing
     */
    public function setPath($pValue = '', $pVerifyFile = true)
    {
        if ($pVerifyFile) {
            if (!file_exists($pValue)) {
                throw new \Exception("File not found : $pValue");
            }

            if ($this->width == 0 && $this->height == 0) {
                // Get width/height
                list($this->width, $this->height) = getimagesize($pValue);
            }
        }
        $this->path = $pValue;
        return $this;
    }

    /**
     * Get Filename
     *
     * @return string
     */
    public function getFilename()
    {
        return basename($this->path);
    }

    /**
     * Get Extension
     *
     * @return string
     */
    public function getExtension()
    {
        $exploded = explode('.', basename($this->path));

        return $exploded[count($exploded) - 1];
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename($numSlide)
    {
        return 'background_' . $numSlide . '.' . $this->getExtension();
    }
}
