<?php

namespace PhpOffice\PhpPresentation\Writer\ODPresentation;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
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
            if ($shape instanceof Drawing) {
                if (!in_array(md5($shape->getPath()), $arrMedia)) {
                    $arrMedia[] = md5($shape->getPath());

                    $imagePath = $shape->getPath();

                    $imageContents = file_get_contents($imagePath);
                    if (strpos($imagePath, 'zip://') !== false) {
                        $imagePath = substr($imagePath, 6);
                        $imagePathSplitted = explode('#', $imagePath);

                        $imageZip = new \ZipArchive();
                        $imageZip->open($imagePathSplitted[0]);
                        $imageContents = $imageZip->getFromName($imagePathSplitted[1]);
                        $imageZip->close();
                        unset($imageZip);
                    } elseif (strpos($imagePath, 'data:image/') === 0) {
                        list(, $imageContents) = explode(';', $imagePath);
                        list(, $imageContents) = explode(',', $imageContents);
                        $imageContents = base64_decode($imageContents);
                    }

                    $this->getZip()->addFromString('Pictures/' . md5($shape->getPath()).'.'.$shape->getExtension(), $imageContents);
                }
            } elseif ($shape instanceof MemoryDrawing) {
                if (!in_array(str_replace(' ', '_', $shape->getIndexedFilename()), $arrMedia)) {
                    $arrMedia[] = str_replace(' ', '_', $shape->getIndexedFilename());
                    ob_start();
                    call_user_func($shape->getRenderingFunction(), $shape->getImageResource());
                    $imageContents = ob_get_contents();
                    ob_end_clean();

                    $this->getZip()->addFromString('Pictures/' . str_replace(' ', '_', $shape->getIndexedFilename()), $imageContents);
                }
            }
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
