<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Slide;
use PhpOffice\Common\XMLWriter;

class PptTheme extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     * @throws \Exception
     */
    public function render()
    {
        foreach ($this->oPresentation->getAllMasterSlides() as $oMasterSlide) {
            $this->getZip()->addFromString('ppt/theme/theme' . $oMasterSlide->getRelsIndex() . '.xml', $this->writeTheme($oMasterSlide));
        }

        return $this->getZip();
    }


    /**
     * Write theme to XML format
     *
     * @param  Slide\SlideMaster $oMasterSlide
     * @return string XML Output
     */
    protected function writeTheme(Slide\SlideMaster $oMasterSlide)
    {
        $arrayFont = array(
            'Jpan' => 'ＭＳ Ｐゴシック',
            'Hang' => '맑은 고딕',
            'Hans' => '宋体',
            'Hant' => '新細明體',
            'Arab' => 'Times New Roman',
            'Hebr' => 'Times New Roman',
            'Thai' => 'Angsana New',
            'Ethi' => 'Nyala',
            'Beng' => 'Vrinda',
            'Gujr' => 'Shruti',
            'Khmr' => 'MoolBoran',
            'Knda' => 'Tunga',
            'Guru' => 'Raavi',
            'Cans' => 'Euphemia',
            'Cher' => 'Plantagenet Cherokee',
            'Yiii' => 'Microsoft Yi Baiti',
            'Tibt' => 'Microsoft Himalaya',
            'Thaa' => 'MV Boli',
            'Deva' => 'Mangal',
            'Telu' => 'Gautami',
            'Taml' => 'Latha',
            'Syrc' => 'Estrangelo Edessa',
            'Orya' => 'Kalinga',
            'Mlym' => 'Kartika',
            'Laoo' => 'DokChampa',
            'Sinh' => 'Iskoola Pota',
            'Mong' => 'Mongolian Baiti',
            'Viet' => 'Times New Roman',
            'Uigh' => 'Microsoft Uighur',
        );

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        $name = 'Theme'.rand(1, 100);

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // a:theme
        $objWriter->startElement('a:theme');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('name', $name);

        // a:theme/a:themeElements
        $objWriter->startElement('a:themeElements');

        // a:theme/a:themeElements/a:clrScheme
        $objWriter->startElement('a:clrScheme');
        $objWriter->writeAttribute('name', $name);

        foreach ($oMasterSlide->getAllSchemeColors() as $oSchemeColor) {
            // a:theme/a:themeElements/a:clrScheme/a:*
            $objWriter->startElement('a:'.$oSchemeColor->getValue());

            if (in_array($oSchemeColor->getValue(), array(
                'dk1', 'lt1'
            ))) {
                $objWriter->startElement('a:sysClr');
                $objWriter->writeAttribute('val', ($oSchemeColor->getValue() == 'dk1' ? 'windowText' : 'window'));
                $objWriter->writeAttribute('lastClr', $oSchemeColor->getRGB());
                $objWriter->endElement();
            } else {
                $objWriter->startElement('a:srgbClr');
                $objWriter->writeAttribute('val', $oSchemeColor->getRGB());
                $objWriter->endElement();
            }

            // a:theme/a:themeElements/a:clrScheme/a:*/
            $objWriter->endElement();
        }

        // a:theme/a:themeElements/a:clrScheme/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fontScheme
        $objWriter->startElement('a:fontScheme');
        $objWriter->writeAttribute('name', $name);

        // a:theme/a:themeElements/a:fontScheme/a:majorFont
        $objWriter->startElement('a:majorFont');

        // a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin
        $objWriter->startElement('a:latin');
        $objWriter->writeAttribute('typeface', 'Calibri');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fontScheme/a:majorFont/a:ea
        $objWriter->startElement('a:ea');
        $objWriter->writeAttribute('typeface', '');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fontScheme/a:majorFont/a:cs
        $objWriter->startElement('a:cs');
        $objWriter->writeAttribute('typeface', '');
        $objWriter->endElement();

        foreach ($arrayFont as $script => $typeface) {
            // a:theme/a:themeElements/a:fontScheme/a:majorFont/a:font
            $objWriter->startElement('a:font');
            $objWriter->writeAttribute('script', $script);
            $objWriter->writeAttribute('typeface', $typeface);
            $objWriter->endElement();
        }

        // a:theme/a:themeElements/a:fontScheme/a:majorFont/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fontScheme/a:minorFont
        $objWriter->startElement('a:minorFont');

        // a:theme/a:themeElements/a:fontScheme/a:majorFont/a:latin
        $objWriter->startElement('a:latin');
        $objWriter->writeAttribute('typeface', 'Calibri');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fontScheme/a:majorFont/a:ea
        $objWriter->startElement('a:ea');
        $objWriter->writeAttribute('typeface', '');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fontScheme/a:majorFont/a:cs
        $objWriter->startElement('a:cs');
        $objWriter->writeAttribute('typeface', '');
        $objWriter->endElement();

        foreach ($arrayFont as $script => $typeface) {
            // a:theme/a:themeElements/a:fontScheme/a:majorFont/a:font
            $objWriter->startElement('a:font');
            $objWriter->writeAttribute('script', $script);
            $objWriter->writeAttribute('typeface', $typeface);
            $objWriter->endElement();
        }

        // a:theme/a:themeElements/a:fontScheme/a:minorFont/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fontScheme/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme
        $objWriter->startElement('a:fmtScheme');
        $objWriter->writeAttribute('name', $name);

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst
        $objWriter->startElement('a:fillStyleLst');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:solidFill
        $objWriter->startElement('a:solidFill');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:solidFill/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:solidFill/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:solidFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill
        $objWriter->startElement('a:gradFill');
        $objWriter->writeAttribute('rotWithShape', 1);

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst
        $objWriter->startElement('a:gsLst');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '0%');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '50%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '300%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '35%');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '37%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '300%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100%');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '15%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '350%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:lin
        $objWriter->startElement('a:lin');
        $objWriter->writeAttribute('ang', 16200000);
        $objWriter->writeAttribute('scaled', 1);
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill
        $objWriter->startElement('a:gradFill');
        $objWriter->writeAttribute('rotWithShape', 1);

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst
        $objWriter->startElement('a:gsLst');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '0%');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '51%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '130%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '80%');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '93%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '130%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100%');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '94%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '135%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:lin
        $objWriter->startElement('a:lin');
        $objWriter->writeAttribute('ang', 16200000);
        $objWriter->writeAttribute('scaled', 0);
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst
        $objWriter->startElement('a:lnStyleLst');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln
        $objWriter->startElement('a:ln');
        $objWriter->writeAttribute('w', 9525);
        $objWriter->writeAttribute('cap', 'flat');
        $objWriter->writeAttribute('cmpd', 'sng');
        $objWriter->writeAttribute('algn', 'ctr');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill
        $objWriter->startElement('a:solidFill');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '95%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '105%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:prstDash
        $objWriter->startElement('a:prstDash');
        $objWriter->writeAttribute('val', 'solid');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln
        $objWriter->startElement('a:ln');
        $objWriter->writeAttribute('w', 25400);
        $objWriter->writeAttribute('cap', 'flat');
        $objWriter->writeAttribute('cmpd', 'sng');
        $objWriter->writeAttribute('algn', 'ctr');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill
        $objWriter->startElement('a:solidFill');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:prstDash
        $objWriter->startElement('a:prstDash');
        $objWriter->writeAttribute('val', 'solid');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln
        $objWriter->startElement('a:ln');
        $objWriter->writeAttribute('w', 38100);
        $objWriter->writeAttribute('cap', 'flat');
        $objWriter->writeAttribute('cmpd', 'sng');
        $objWriter->writeAttribute('algn', 'ctr');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill
        $objWriter->startElement('a:solidFill');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:solidFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/a:prstDash
        $objWriter->startElement('a:prstDash');
        $objWriter->writeAttribute('val', 'solid');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/a:ln/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:lnStyleLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst
        $objWriter->startElement('a:effectStyleLst');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle
        $objWriter->startElement('a:effectStyle');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst
        $objWriter->startElement('a:effectLst');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw
        $objWriter->startElement('a:outerShdw');
        $objWriter->writeAttribute('blurRad', 40000);
        $objWriter->writeAttribute('dir', 5400000);
        $objWriter->writeAttribute('dist', 20000);
        $objWriter->writeAttribute('rotWithShape', 0);

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', '000000');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/a:alpha
        $objWriter->startElement('a:alpha');
        $objWriter->writeAttribute('val', '38%');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/a:alpha/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle
        $objWriter->startElement('a:effectStyle');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst
        $objWriter->startElement('a:effectLst');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw
        $objWriter->startElement('a:outerShdw');
        $objWriter->writeAttribute('blurRad', 40000);
        $objWriter->writeAttribute('dir', 5400000);
        $objWriter->writeAttribute('dist', 23000);
        $objWriter->writeAttribute('rotWithShape', 0);

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', '000000');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/a:alpha
        $objWriter->startElement('a:alpha');
        $objWriter->writeAttribute('val', '35%');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/a:alpha/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle
        $objWriter->startElement('a:effectStyle');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst
        $objWriter->startElement('a:effectLst');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw
        $objWriter->startElement('a:outerShdw');
        $objWriter->writeAttribute('blurRad', 40000);
        $objWriter->writeAttribute('dir', 5400000);
        $objWriter->writeAttribute('dist', 23000);
        $objWriter->writeAttribute('rotWithShape', 0);

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr
        $objWriter->startElement('a:srgbClr');
        $objWriter->writeAttribute('val', '000000');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/a:alpha
        $objWriter->startElement('a:alpha');
        $objWriter->writeAttribute('val', '35%');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/a:alpha/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/a:srgbClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/a:outerShdw/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:effectLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:scene3d
        $objWriter->startElement('a:scene3d');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:scene3d/a:camera
        $objWriter->startElement('a:camera');
        $objWriter->writeAttribute('prst', 'orthographicFront');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:scene3d/a:camera/a:rot
        $objWriter->startElement('a:rot');
        $objWriter->writeAttribute('lat', 0);
        $objWriter->writeAttribute('lon', 0);
        $objWriter->writeAttribute('rev', 0);
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:scene3d/a:camera/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:scene3d/a:lightRig
        $objWriter->startElement('a:lightRig');
        $objWriter->writeAttribute('dir', 't');
        $objWriter->writeAttribute('rig', 'threePt');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:scene3d/a:lightRig/a:rot
        $objWriter->startElement('a:rot');
        $objWriter->writeAttribute('lat', 0);
        $objWriter->writeAttribute('lon', 0);
        $objWriter->writeAttribute('rev', 1200000);
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:scene3d/a:lightRig/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:scene3d/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:sp3d
        $objWriter->startElement('a:sp3d');

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:sp3d/a:bevelT
        $objWriter->startElement('a:bevelT');
        $objWriter->writeAttribute('h', 25400);
        $objWriter->writeAttribute('w', 63500);
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/a:sp3d/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/a:effectStyle/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:effectStyleLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst
        $objWriter->startElement('a:bgFillStyleLst');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill
        $objWriter->startElement('a:solidFill');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:solidFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill
        $objWriter->startElement('a:gradFill');
        $objWriter->writeAttribute('rotWithShape', 1);

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst
        $objWriter->startElement('a:gsLst');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '0%');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '40%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '350%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '40%');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '45%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '99%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '350%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100%');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '20%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '255%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:path
        $objWriter->startElement('a:path');
        $objWriter->writeAttribute('path', 'circle');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:path/a:fillToRect
        $objWriter->startElement('a:fillToRect');
        $objWriter->writeAttribute('b', '180%');
        $objWriter->writeAttribute('l', '50%');
        $objWriter->writeAttribute('r', '50%');
        $objWriter->writeAttribute('t', '-80%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:path/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill
        $objWriter->startElement('a:gradFill');
        $objWriter->writeAttribute('rotWithShape', 1);

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst
        $objWriter->startElement('a:gsLst');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '0%');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '80%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '300%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100%');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '30%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '200%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:path
        $objWriter->startElement('a:path');
        $objWriter->writeAttribute('path', 'circle');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:path/a:fillToRect
        $objWriter->startElement('a:fillToRect');
        $objWriter->writeAttribute('b', '50%');
        $objWriter->writeAttribute('l', '50%');
        $objWriter->writeAttribute('r', '50%');
        $objWriter->writeAttribute('t', '50%');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:path/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/
        $objWriter->endElement();

        // a:theme/a:themeElements/
        $objWriter->endElement();

        // a:theme/a:themeElements
        $objWriter->writeElement('a:objectDefaults');

        // a:theme/a:extraClrSchemeLst
        $objWriter->writeElement('a:extraClrSchemeLst');

        // a:theme/
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}
