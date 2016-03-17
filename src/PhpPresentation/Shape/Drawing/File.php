<?php

namespace PhpOffice\PhpPresentation\Shape\Drawing;

use SplFileInfo;

class File extends AbstractDrawingAdapter
{
    /**
     * @var SplFileInfo
     */
    protected $oFileInfo;

    /**
     * Get Path
     *
     * @return string
     */
    public function getPath()
    {
        return $this->oFileInfo->getPathname();
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
                throw new \Exception("File $pValue not found!");
            }
        }
        $this->oFileInfo = new SplFileInfo($pValue);

        if ($this->width == 0 && $this->height == 0) {
            list($this->width, $this->height) = getimagesize($this->getPath());
        }

        return $this;
    }

    /**
     * @return string
     */
    public function getContents()
    {
        return file_get_contents($this->getPath());
    }


    /**
     * @return string
     */
    public function getExtension()
    {
        return $this->oFileInfo->getExtension();
    }

    /**
     * @return string
     */
    public function getMimeType()
    {
        $image = getimagesize($this->getPath());
        return image_type_to_mime_type($image[2]);
    }

    /**
     * @return string
     */
    public function getIndexedFilename()
    {
        $output = str_replace('.' . $this->getExtension(), '', $this->oFileInfo->getFilename());
        $output .= $this->getImageIndex();
        $output .= $this->getExtension();
        $output = str_replace(' ', '_', $output);
        return $output;
    }
}