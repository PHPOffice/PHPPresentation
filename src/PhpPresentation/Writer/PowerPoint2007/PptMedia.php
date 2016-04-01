<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Drawing;

class PptMedia extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
            $shape = $this->getDrawingHashTable()->getByIndex($i);
            if (!$shape instanceof Drawing\AbstractDrawingAdapter) {
                continue;
            }
            $this->getZip()->addFromString('ppt/media/' . $shape->getIndexedFilename(), $shape->getContents());
        }

        return $this->getZip();
    }
}
