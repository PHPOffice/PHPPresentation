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

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Slide;

class PptTheme extends AbstractDecoratorWriter
{
    public function render(): ZipInterface
    {
        foreach ($this->oPresentation->getAllMasterSlides() as $oMasterSlide) {
            $this->getZip()->addFromString('ppt/theme/theme' . $oMasterSlide->getRelsIndex() . '.xml', $this->writeTheme($oMasterSlide));
        }

        return $this->getZip();
    }

    /**
     * Write theme to XML format.
     *
     * @return string XML Output
     */
    protected function writeTheme(Slide\SlideMaster $oMasterSlide): string
    {
        $arrayFont = [
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
        ];

        // Create XML writer
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);

        $name = 'Theme' . mt_rand(1, 100);

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
            $objWriter->startElement('a:' . $oSchemeColor->getValue());

            if (in_array($oSchemeColor->getValue(), [
                'dk1', 'lt1',
            ])) {
                $objWriter->startElement('a:sysClr');
                $objWriter->writeAttribute('val', ('dk1' == $oSchemeColor->getValue() ? 'windowText' : 'window'));
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
        $objWriter->writeAttribute('pos', '0');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '50000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '300000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '35000');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '37000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '300000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100000');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '15000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '350000');
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
        $objWriter->writeAttribute('pos', '0');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '51000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '130000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '80000');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '93000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '130000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100000');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '94000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '135000');
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
        $objWriter->writeAttribute('val', '95000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:fillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '105000');
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
        $objWriter->writeAttribute('val', '38000');

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
        $objWriter->writeAttribute('val', '35');

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
        $objWriter->writeAttribute('val', '35000');

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
        $objWriter->writeAttribute('pos', '0');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '40000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '350000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '40000');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '45000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '99000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '350000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100000');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '20000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '255000');
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
        $objWriter->writeAttribute('b', '180000');
        $objWriter->writeAttribute('l', '50000');
        $objWriter->writeAttribute('r', '50000');
        $objWriter->writeAttribute('t', '-80000');
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
        $objWriter->writeAttribute('pos', '0');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:tint
        $objWriter->startElement('a:tint');
        $objWriter->writeAttribute('val', '80000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '300000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs
        $objWriter->startElement('a:gs');
        $objWriter->writeAttribute('pos', '100000');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr
        $objWriter->startElement('a:schemeClr');
        $objWriter->writeAttribute('val', 'phClr');

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:shade
        $objWriter->startElement('a:shade');
        $objWriter->writeAttribute('val', '30000');
        $objWriter->endElement();

        // a:theme/a:themeElements/a:fmtScheme/a:bgFillStyleLst/a:gradFill/a:gsLst/a:gs/a:schemeClr/a:satMod
        $objWriter->startElement('a:satMod');
        $objWriter->writeAttribute('val', '200000');
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
        $objWriter->writeAttribute('b', '50000');
        $objWriter->writeAttribute('l', '50000');
        $objWriter->writeAttribute('r', '50000');
        $objWriter->writeAttribute('t', '50000');
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
