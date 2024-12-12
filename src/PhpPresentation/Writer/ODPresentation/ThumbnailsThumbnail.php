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
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;

class ThumbnailsThumbnail extends AbstractDecoratorWriter
{
    public function render(): ZipInterface
    {
        $pathThumbnail = $this->getPresentation()->getPresentationProperties()->getThumbnailPath();
        if ($pathThumbnail) {
            // Size : 128x128 pixel
            // PNG : 8bit, non-interlaced with full alpha transparency
            $gdImage = imagecreatefromstring(file_get_contents($pathThumbnail));
            if ($gdImage) {
                [$width, $height] = getimagesize($pathThumbnail);

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
