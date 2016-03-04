<?php
namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\XMLWriter;

class PptTableProps extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     */
    public function render()
    {
        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // a:tblStyleLst
        $objWriter->startElement('a:tblStyleLst');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('def', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
        $objWriter->endElement();

        $this->getZip()->addFromString('ppt/tableStyles.xml', $objWriter->getData());

        return $this->getZip();
    }
}
