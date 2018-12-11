<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Slide;

class PptComments extends AbstractDecoratorWriter
{
    /**
     * @return ZipInterface
     */
    public function render()
    {
        foreach ($this->getPresentation()->getAllSlides() as $numSlide => $oSlide) {
            $contentXml = $this->writeSlideComments($oSlide);
            if (empty($contentXml)) {
                continue;
            }
            $this->getZip()->addFromString('ppt/comments/comment'.($numSlide + 1).'.xml', $contentXml);
        }
        return $this->getZip();
    }

    /**
     * @param Slide $oSlide
     * @return string
     */
    protected function writeSlideComments(Slide $oSlide)
    {
        /**
         * @var Comment[]
         */
        $arrayComment = array();
        foreach ($oSlide->getShapeCollection() as $oShape) {
            if ($oShape instanceof Comment) {
                $arrayComment[] = $oShape;
            }
        }

        if (empty($arrayComment)) {
            return '';
        }

        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:cmLst
        $objWriter->startElement('p:cmLst');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        foreach ($arrayComment as $idxComment => $oComment) {
            $oAuthor = $oComment->getAuthor();

            // p:cmLst > p:cm
            $objWriter->startElement('p:cm');
            $objWriter->writeAttribute('authorId', $oAuthor instanceof Comment\Author ? $oAuthor->getIndex() : 0);
            $objWriter->writeAttribute('dt', date('Y-m-d\TH:i:s.000000000', $oComment->getDate()));
            $objWriter->writeAttribute('idx', $idxComment);

            // p:cmLst > p:cm > p:pos
            // Uses 1/8pt for positionning
            // @link : https://social.msdn.microsoft.com/Forums/fr-FR/ebdc12f2-0cff-4fa8-b901-fa6e3198364e/comment-position-units
            $objWriter->startElement('p:pos');
            $objWriter->writeAttribute('x', (int) CommonDrawing::pixelsToPoints($oComment->getOffsetX() * 8));
            $objWriter->writeAttribute('y', (int) CommonDrawing::pixelsToPoints($oComment->getOffsetY() * 8));
            $objWriter->endElement();

            // p:cmLst > p:cm > p:text
            $objWriter->writeElement('p:text', $oComment->getText());

            // p:cmLst > ## p:cm
            $objWriter->endElement();
        }

        // ## p:cmLst
        $objWriter->endElement();

        return $objWriter->getData();
    }
}
