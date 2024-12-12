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
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Slide\Background\Image;

class Pictures extends AbstractDecoratorWriter
{
    public function render(): ZipInterface
    {
        $arrMedia = [];
        for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
            $shape = $this->getDrawingHashTable()->getByIndex($i);
            if (!($shape instanceof Drawing\AbstractDrawingAdapter)) {
                continue;
            }
            $arrMedia[] = $shape->getIndexedFilename();
            $this->getZip()->addFromString('Pictures/' . $shape->getIndexedFilename(), $shape->getContents());
        }

        foreach ($this->getPresentation()->getAllSlides() as $keySlide => $oSlide) {
            // Add background image slide
            $oBkgImage = $oSlide->getBackground();
            if ($oBkgImage instanceof Image) {
                $this->getZip()->addFromString('Pictures/' . $oBkgImage->getIndexedFilename((string) $keySlide), file_get_contents($oBkgImage->getPath()));
            }
        }

        return $this->getZip();
    }
}
