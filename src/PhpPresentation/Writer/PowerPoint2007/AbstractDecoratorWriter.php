<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;

abstract class AbstractDecoratorWriter
{
    /**
     * @return ZipInterface
     */
    abstract public function render();


    /**
     * Write relationship
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter   XML Writer
     * @param  int                            $pId         Relationship ID. rId will be prepended!
     * @param  string                         $pType       Relationship type
     * @param  string                         $pTarget     Relationship target
     * @param  string                         $pTargetMode Relationship target mode
     * @throws \Exception
     */
    protected function writeRelationship(XMLWriter $objWriter, $pId = 1, $pType = '', $pTarget = '', $pTargetMode = '')
    {
        if ($pType != '' && $pTarget != '') {
            if (strpos($pId, 'rId') === false) {
                $pId = 'rId' . $pId;
            }

            // Write relationship
            $objWriter->startElement('Relationship');
            $objWriter->writeAttribute('Id', $pId);
            $objWriter->writeAttribute('Type', $pType);
            $objWriter->writeAttribute('Target', $pTarget);

            if ($pTargetMode != '') {
                $objWriter->writeAttribute('TargetMode', $pTargetMode);
            }

            $objWriter->endElement();
        } else {
            throw new \Exception("Invalid parameters passed.");
        }
    }

    /**
     * @var \PhpOffice\PhpPresentation\HashTable
     */
    protected $hashTable;

    public function setDrawingHashTable(HashTable $hashTable)
    {
        $this->hashTable = $hashTable;
        return $this;
    }

    /**
     * @return HashTable
     */
    public function getDrawingHashTable()
    {
        return $this->hashTable;
    }

    /**
     * @var ZipInterface
     */
    protected $oZip;

    public function setZip(ZipInterface $oZip)
    {
        $this->oZip = $oZip;
        return $this;
    }

    /**
     * @return ZipInterface
     */
    public function getZip()
    {
        return $this->oZip;
    }

    /**
     * @var PhpPresentation
     */
    protected $oPresentation;

    public function setPresentation(PhpPresentation $oPresentation)
    {
        $this->oPresentation = $oPresentation;
        return $this;
    }

    /**
     * @return PhpPresentation
     */
    public function getPresentation()
    {
        return $this->oPresentation;
    }

    /**
     * Write Border
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter    XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Border     $pBorder      Border
     * @param  string                         $pElementName Element name
     * @throws \Exception
     */
    protected function writeBorder(XMLWriter $objWriter, Border $pBorder, $pElementName = 'L')
    {
        // Line style
        $lineStyle = $pBorder->getLineStyle();
        if ($lineStyle == Border::LINE_NONE) {
            $lineStyle = Border::LINE_SINGLE;
        }

        // Line width
        $lineWidth = 12700 * $pBorder->getLineWidth();

        // a:ln $pElementName
        $objWriter->startElement('a:ln' . $pElementName);
        $objWriter->writeAttribute('w', $lineWidth);
        $objWriter->writeAttribute('cap', 'flat');
        $objWriter->writeAttribute('cmpd', $lineStyle);
        $objWriter->writeAttribute('algn', 'ctr');

        // Fill?
        if ($pBorder->getLineStyle() == Border::LINE_NONE) {
            // a:noFill
            $objWriter->writeElement('a:noFill', null);
        } else {
            // a:solidFill
            $objWriter->startElement('a:solidFill');

            // a:srgbClr
            $objWriter->startElement('a:srgbClr');
            $objWriter->writeAttribute('val', $pBorder->getColor()->getRGB());
            $objWriter->endElement();

            $objWriter->endElement();
        }

        // Dash
        $objWriter->startElement('a:prstDash');
        $objWriter->writeAttribute('val', $pBorder->getDashStyle());
        $objWriter->endElement();

        // a:round
        $objWriter->writeElement('a:round', null);

        // a:headEnd
        $objWriter->startElement('a:headEnd');
        $objWriter->writeAttribute('type', 'none');
        $objWriter->writeAttribute('w', 'med');
        $objWriter->writeAttribute('len', 'med');
        $objWriter->endElement();

        // a:tailEnd
        $objWriter->startElement('a:tailEnd');
        $objWriter->writeAttribute('type', 'none');
        $objWriter->writeAttribute('w', 'med');
        $objWriter->writeAttribute('len', 'med');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Fill
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Fill       $pFill     Fill style
     * @throws \Exception
     */
    protected function writeFill(XMLWriter $objWriter, Fill $pFill)
    {
        // Is it a fill?
        if ($pFill->getFillType() == Fill::FILL_NONE) {
            return;
        }

        // Is it a solid fill?
        if ($pFill->getFillType() == Fill::FILL_SOLID) {
            $this->writeSolidFill($objWriter, $pFill);
            return;
        }

        // Check if this is a pattern type or gradient type
        if ($pFill->getFillType() == Fill::FILL_GRADIENT_LINEAR || $pFill->getFillType() == Fill::FILL_GRADIENT_PATH) {
            // Gradient fill
            $this->writeGradientFill($objWriter, $pFill);
        } else {
            // Pattern fill
            $this->writePatternFill($objWriter, $pFill);
        }
    }


    /**
     * Write Solid Fill
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Fill       $pFill     Fill style
     * @throws \Exception
     */
    protected function writeSolidFill(XMLWriter $objWriter, Fill $pFill)
    {
        // a:gradFill
        $objWriter->startElement('a:solidFill');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getStartColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Gradient Fill
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Fill       $pFill     Fill style
     * @throws \Exception
     */
    protected function writeGradientFill(XMLWriter $objWriter, Fill $pFill)
    {
        // a:gradFill
        $objWriter->startElement('a:gradFill');

        // a:gsLst
        $objWriter->startElement('a:gsLst');
        // a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '0');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getStartColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();

        // a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100000');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getEndColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();

        // a:lin
        $objWriter->startElement('a:lin');
        $objWriter->writeAttribute('ang', CommonDrawing::degreesToAngle($pFill->getRotation()));
        $objWriter->writeAttribute('scaled', '0');
        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Write Pattern Fill
     *
     * @param  \PhpOffice\Common\XMLWriter $objWriter XML Writer
     * @param  \PhpOffice\PhpPresentation\Style\Fill       $pFill     Fill style
     * @throws \Exception
     */
    protected function writePatternFill(XMLWriter $objWriter, Fill $pFill)
    {
        // a:pattFill
        $objWriter->startElement('a:pattFill');

        // fgClr
        $objWriter->startElement('a:fgClr');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getStartColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();

        // bgClr
        $objWriter->startElement('a:bgClr');

        // srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', $pFill->getEndColor()->getRGB());
        $objWriter->endElement();

        $objWriter->endElement();

        $objWriter->endElement();
    }

    /**
     * Determine absolute zip path
     *
     * @param  string $path
     * @return string
     */
    protected function absoluteZipPath($path)
    {
        $path      = str_replace(array(
            '/',
            '\\'
        ), DIRECTORY_SEPARATOR, $path);
        $parts     = array_filter(explode(DIRECTORY_SEPARATOR, $path), 'strlen');
        $absolutes = array();
        foreach ($parts as $part) {
            if ('.' == $part) {
                continue;
            }
            if ('..' == $part) {
                array_pop($absolutes);
            } else {
                $absolutes[] = $part;
            }
        }

        return implode('/', $absolutes);
    }
}
