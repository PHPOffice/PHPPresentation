<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Slide\Background\Image;

class Pictures extends AbstractDecoratorWriter
{
    /**
     * @return ZipInterface
     */

    public function render()
    {
        $arrMedia = array();
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
                $this->getZip()->addFromString('Pictures/'.$oBkgImage->getIndexedFilename($keySlide), file_get_contents($oBkgImage->getPath()));
            }
        }
        
        return $this->getZip();
    }
}
