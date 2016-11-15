<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

class DocPropsThumbnail extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        $pathThumbnail = $this->getPresentation()->getPresentationProperties()->getThumbnailPath();

        if ($pathThumbnail) {
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

        // Return
        return $this->getZip();
    }
}
