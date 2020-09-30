<?php

namespace PhpOffice\PhpPresentation\Shape\Drawing;

use PhpOffice\Common\File as CommonFile;

class File extends AbstractDrawingAdapter
{
    /**
     * @var string
     */
    protected $path;

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
     * @return \PhpOffice\PhpPresentation\Shape\Drawing\File
     */
    public function setPath($pValue = '', $pVerifyFile = true)
    {
        if ($pVerifyFile) {
            if (!file_exists($pValue)) {
                throw new \Exception("File $pValue not found!");
            }
        }
        $this->path = $pValue;

        if ($pVerifyFile) {
            if ($this->width == 0 && $this->height == 0) {
                list($this->width, $this->height) = getimagesize($this->getPath());
            }
        }

        return $this;
    }

    /**
     * @return string
     */
    public function getContents()
    {
        return CommonFile::fileGetContents($this->getPath());
    }


    /**
     * @return string
     */
    public function getExtension()
    {
        return pathinfo($this->getPath(), PATHINFO_EXTENSION);
    }

    /**
     * @throws \Exception
     * @return string
     */
    public function getMimeType()
    {
        if (!CommonFile::fileExists($this->getPath())) {
            throw new \Exception('File '.$this->getPath().' does not exist');
        }
        $image = getimagesizefromstring(CommonFile::fileGetContents($this->getPath()));
        return image_type_to_mime_type($image[2]);
    }

    /**
     * @return string
     */
    public function getIndexedFilename()
    {
        $output = str_replace('.' . $this->getExtension(), '', pathinfo($this->getPath(), PATHINFO_FILENAME));
        $output .= $this->getImageIndex();
        $output .= '.'.$this->getExtension();
        $output = str_replace(' ', '_', $output);
        return $output;
    }
}
