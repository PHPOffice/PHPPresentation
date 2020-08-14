<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

class DocPropsThumbnail extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     * @throws \Exception
     */
    public function render()
    {
        $pathThumbnail = $this->getPresentation()->getPresentationProperties()->getThumbnailPath();
        $type = $this->getPresentation()->getPresentationProperties()->getThumbnailType();

        // From local file
        if ($pathThumbnail && $type == \PhpOffice\PhpPresentation\PresentationProperties::THUMBNAIL_FILE) {
            $fileThumbnail = file_get_contents($pathThumbnail);
            $gdImage = imagecreatefromstring($fileThumbnail);
            if ($gdImage) {
                ob_start();
                imagejpeg($gdImage);
                $imageContents = ob_get_contents();
                ob_end_clean();
                imagedestroy($gdImage);
                $this->getZip()->addFromString('docProps/thumbnail.jpeg', $imageContents);
            }
        }
  
        // From ZIP original file
        if ($pathThumbnail && $type == \PhpOffice\PhpPresentation\PresentationProperties::THUMBNAIL_ZIP) {
            $gdImage = imagecreatefromstring($this->getPresentation()->getPresentationProperties()->getThumbnail());
            if ($gdImage) {
                ob_start();
                imagejpeg($gdImage);
                $imageContents = ob_get_contents();
                ob_end_clean();
                imagedestroy($gdImage);
                $this->getZip()->addFromString('docProps/thumbnail.jpeg', $imageContents);
            }
        }

        // Return
        return $this->getZip();
    }
}
