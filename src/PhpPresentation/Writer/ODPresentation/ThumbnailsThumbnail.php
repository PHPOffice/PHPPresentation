<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

class ThumbnailsThumbnail extends AbstractDecoratorWriter
{
    /**
     * @return ZipInterface
     */
    public function render()
    {
        $pathThumbnail = $this->getPresentation()->getPresentationProperties()->getThumbnailPath();
        if ($pathThumbnail) {
            // Size : 128x128 pixel
            // PNG : 8bit, non-interlaced with full alpha transparency
            $gdImage = imagecreatefromstring(file_get_contents($pathThumbnail));
            if ($gdImage) {
                list($width, $height) = getimagesize($pathThumbnail);

                $gdRender = imagecreatetruecolor(128, 128);
                $colorBgAlpha = imagecolorallocatealpha($gdRender, 0, 0, 0, 127);
                imagecolortransparent($gdRender, $colorBgAlpha);
                imagefill($gdRender, 0, 0, $colorBgAlpha);
                imagecopyresampled($gdRender, $gdImage, 0, 0, 0, 0, 128, 128, $width, $height);
                imagetruecolortopalette($gdRender, false, 255);
                imagesavealpha($gdRender, true);

                ob_start();
                imagepng($gdRender);
                $imageContents = ob_get_contents();
                ob_end_clean();

                imagedestroy($gdRender);
                imagedestroy($gdImage);

                $this->getZip()->addFromString('Thumbnails/thumbnail.png', $imageContents);
            }
        }
        return $this->getZip();
    }
}
