<?php

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PhpOffice\PhpPresentation\Writer\ODPresentation;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\Writer\ODPresentation
 */
class StylesTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    public function testDocumentLayout()
    {
        $element = "/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties";

        $oDocumentLayout = new DocumentLayout();
        $oDocumentLayout->setDocumentLayout(DocumentLayout::LAYOUT_A4, true);
        $this->oPresentation->setLayout($oDocumentLayout);

        $this->assertZipXmlElementExists('styles.xml', $element);
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'style:print-orientation', 'landscape');
        
        $oDocumentLayout->setDocumentLayout(DocumentLayout::LAYOUT_A4, false);
        $this->oPresentation->setLayout($oDocumentLayout);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('styles.xml', $element);
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'style:print-orientation', 'portrait');
    }
    
    public function testCustomDocumentLayout()
    {
        $oDocumentLayout = new DocumentLayout();
        $oDocumentLayout->setDocumentLayout(array('cx' => rand(1, 100),'cy' => rand(1, 100),));
        $this->oPresentation->setLayout($oDocumentLayout);
        
        $element = "/office:document-styles/office:automatic-styles/style:page-layout";
        $this->assertZipXmlElementExists('styles.xml', $element);
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'style:name', 'sPL0');
        
        $element = "/office:document-styles/office:master-styles/style:master-page";
        $this->assertZipXmlElementExists('styles.xml', $element);
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'style:page-layout-name', 'sPL0');
    }
    
    public function testGradientTable()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FFFF7700'))->setEndColor(new Color('FFFFFFFF'));

        $element = "/office:document-styles/office:styles/draw:gradient";
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:name', 'gradient_' . $oCell->getFill()->getHashCode());
    }
    
    public function testStrokeDash()
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setLineStyle(Border::LINE_SINGLE);
        $arrayDashStyle = array(
            Border::DASH_DASH,
            Border::DASH_DASHDOT,
            Border::DASH_DOT,
            Border::DASH_LARGEDASH,
            Border::DASH_LARGEDASHDOT,
            Border::DASH_LARGEDASHDOTDOT,
            Border::DASH_SYSDASH,
            Border::DASH_SYSDASHDOT,
            Border::DASH_SYSDASHDOTDOT,
            Border::DASH_SYSDOT,
        );
        
        foreach ($arrayDashStyle as $style) {
            $oRichText1->getBorder()->setDashStyle($style);

            $element = '/office:document-styles/office:styles/draw:stroke-dash[@draw:name=\'strokeDash_'.$style.'\']';
            $this->assertZipXmlElementExists('styles.xml', $element);
            $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:style', 'rect');
            $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:distance');
            
            switch ($style) {
                case Border::DASH_DOT:
                case Border::DASH_SYSDOT:
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots1');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots1-length');
                    break;
                case Border::DASH_DASH:
                case Border::DASH_LARGEDASH:
                case Border::DASH_SYSDASH:
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots2');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots2-length');
                    break;
                case Border::DASH_DASHDOT:
                case Border::DASH_LARGEDASHDOT:
                case Border::DASH_LARGEDASHDOTDOT:
                case Border::DASH_SYSDASHDOT:
                case Border::DASH_SYSDASHDOTDOT:
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots1');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots1-length');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots2');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots2-length');
                    break;
            }
            $this->resetPresentationFile();
        }
    }
}
