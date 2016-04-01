<?php

namespace PhpOffice\PhpPresentation\Shape\Drawing;

use PhpOffice\Common\File as CommonFile;

class ZipFile extends AbstractDrawingAdapter
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
     * @return \PhpOffice\PhpPresentation\Shape\Drawing
     */
    public function setPath($pValue = '')
    {
        $this->path = $pValue;
        return $this;
    }

    /**
     * @return string
     */
    public function getContents()
    {
        $imagePath         = substr($this->getPath(), 6);
        $imagePathSplitted = explode('#', $imagePath);

        $imageZip = new \ZipArchive();
        $imageZip->open($imagePathSplitted[0]);
        $imageContents = $imageZip->getFromName($imagePathSplitted[1]);
        $imageZip->close();
        unset($imageZip);
        return $imageContents;
    }


    /**
     * @return string
     */
    public function getExtension()
    {
        $imagePath         = substr($this->getPath(), 6);
        $imagePathSplitted = explode('#', $imagePath);
        return pathinfo($imagePathSplitted[1], PATHINFO_EXTENSION);
    }

    /**
     * @return string
     */
    public function getMimeType()
    {
        $pZIPFile = str_replace('zip://', '', $this->getPath());
        $pZIPFile = substr($pZIPFile, 0, strpos($pZIPFile, '#'));
        if (!CommonFile::fileExists($pZIPFile)) {
            throw new \Exception("File $pZIPFile does not exist");
        }
        $pImgFile = substr($this->getPath(), strpos($this->getPath(), '#') + 1);
        $oArchive = new \ZipArchive();
        $oArchive->open($pZIPFile);
        if (!function_exists('getimagesizefromstring')) {
            $uri = 'data://application/octet-stream;base64,' . base64_encode($oArchive->getFromName($pImgFile));
            $image = getimagesize($uri);
        } else {
            $image = getimagesizefromstring($oArchive->getFromName($pImgFile));
        }
        return image_type_to_mime_type($image[2]);
    }

    /**
     * @return string
     */
    public function getIndexedFilename()
    {
        $output = substr($this->getPath(), strpos($this->getPath(), '#') + 1);
        $output = str_replace('.' . $this->getExtension(), '', $output);
        $output .= $this->getImageIndex();
        $output .= '.'.$this->getExtension();
        $output = str_replace(' ', '_', $output);
        return $output;
    }
}