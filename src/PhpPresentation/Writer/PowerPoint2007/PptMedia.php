<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;

class PptMedia extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
            $shape = $this->getDrawingHashTable()->getByIndex($i);
            if ($shape instanceof Drawing) {
                $imagePath     = $shape->getPath();
                if (strpos($imagePath, 'zip://') !== false) {
                    $imagePath         = substr($imagePath, 6);
                    $imagePathSplitted = explode('#', $imagePath);

                    $imageZip = new \ZipArchive();
                    $imageZip->open($imagePathSplitted[0]);
                    $imageContents = $imageZip->getFromName($imagePathSplitted[1]);
                    $imageZip->close();
                    unset($imageZip);
                } else {
                    $imageContents = file_get_contents($imagePath);
                }
                $this->getZip()->addFromString('ppt/media/' . str_replace(' ', '_', $shape->getIndexedFilename()), $imageContents);
            } elseif ($shape instanceof MemoryDrawing) {
                ob_start();
                call_user_func($shape->getRenderingFunction(), $shape->getImageResource());
                $imageContents = ob_get_contents();
                ob_end_clean();

                $this->getZip()->addFromString('ppt/media/' . str_replace(' ', '_', $shape->getIndexedFilename()), $imageContents);
            }
        }

        return $this->getZip();
    }
}
