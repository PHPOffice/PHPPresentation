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
     * Get Path.
     */
    public function getPath(): string
    {
        return $this->path;
    }

    /**
     * Set Path.
     *
     * @param string $pValue File path
     *
     * @return \PhpOffice\PhpPresentation\Shape\Drawing\ZipFile
     */
    public function setPath(string $pValue = ''): self
    {
        $this->path = $pValue;

        return $this;
    }

    /**
     * @throws \Exception
     */
    public function getContents(): string
    {
        if (!CommonFile::fileExists($this->getZipFileOut())) {
            throw new \Exception('File ' . $this->getZipFileOut() . ' does not exist');
        }

        $imageZip = new \ZipArchive();
        $imageZip->open($this->getZipFileOut());
        $imageContents = $imageZip->getFromName($this->getZipFileIn());
        $imageZip->close();
        unset($imageZip);

        return $imageContents;
    }

    public function getExtension(): string
    {
        return pathinfo($this->getZipFileIn(), PATHINFO_EXTENSION);
    }

    /**
     * @throws \Exception
     */
    public function getMimeType(): string
    {
        if (!CommonFile::fileExists($this->getZipFileOut())) {
            throw new \Exception('File ' . $this->getZipFileOut() . ' does not exist');
        }
        $oArchive = new \ZipArchive();
        $oArchive->open($this->getZipFileOut());
        if (!function_exists('getimagesizefromstring')) {
            $uri = 'data://application/octet-stream;base64,' . base64_encode($oArchive->getFromName($this->getZipFileIn()));
            $image = getimagesize($uri);
        } else {
            $image = getimagesizefromstring($oArchive->getFromName($this->getZipFileIn()));
        }

        return image_type_to_mime_type($image[2]);
    }

    public function getIndexedFilename(): string
    {
        $output = pathinfo($this->getZipFileIn(), PATHINFO_FILENAME);
        $output = str_replace('.' . $this->getExtension(), '', $output);
        $output .= $this->getImageIndex();
        $output .= '.' . $this->getExtension();
        $output = str_replace(' ', '_', $output);

        return $output;
    }

    protected function getZipFileOut(): string
    {
        $path = str_replace('zip://', '', $this->getPath());
        $path = explode('#', $path);

        return empty($path[0]) ? '' : $path[0];
    }

    protected function getZipFileIn(): string
    {
        $path = str_replace('zip://', '', $this->getPath());
        $path = explode('#', $path);

        return empty($path[1]) ? '' : $path[1];
    }
}
