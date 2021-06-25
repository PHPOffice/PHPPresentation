<?php

namespace PhpOffice\PhpPresentation\Shape\Drawing;

use PhpOffice\Common\File as CommonFile;

class File extends AbstractDrawingAdapter
{
    /**
     * @var string
     */
    protected $path = '';

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
     * @param bool $pVerifyFile Verify file
     *
     * @throws \Exception
     *
     * @return \PhpOffice\PhpPresentation\Shape\Drawing\File
     */
    public function setPath(string $pValue = '', $pVerifyFile = true): self
    {
        if ($pVerifyFile) {
            if (!file_exists($pValue)) {
                throw new \Exception("File $pValue not found!");
            }
        }
        $this->path = $pValue;

        if ($pVerifyFile) {
            if (0 == $this->width && 0 == $this->height) {
                list($this->width, $this->height) = getimagesize($this->getPath());
            }
        }

        return $this;
    }

    public function getContents(): string
    {
        return CommonFile::fileGetContents($this->getPath());
    }

    public function getExtension(): string
    {
        return pathinfo($this->getPath(), PATHINFO_EXTENSION);
    }

    /**
     * @throws \Exception
     */
    public function getMimeType(): string
    {
        if (!CommonFile::fileExists($this->getPath())) {
            throw new \Exception('File ' . $this->getPath() . ' does not exist');
        }
        $image = getimagesizefromstring(CommonFile::fileGetContents($this->getPath()));

        return image_type_to_mime_type($image[2]);
    }

    public function getIndexedFilename(): string
    {
        $output = str_replace('.' . $this->getExtension(), '', pathinfo($this->getPath(), PATHINFO_FILENAME));
        $output .= $this->getImageIndex();
        $output .= '.' . $this->getExtension();
        $output = str_replace(' ', '_', $output);

        return $output;
    }
}
