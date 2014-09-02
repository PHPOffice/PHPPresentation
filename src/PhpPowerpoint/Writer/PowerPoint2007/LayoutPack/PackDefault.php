<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\LayoutPack;

use PhpOffice\PhpPowerpoint\Slide\Layout;
use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\AbstractLayoutPack;

/**
 * \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\LayoutPack\Default
 */
class PackDefault extends AbstractLayoutPack
{
    /**
     * \PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\LayoutPack\Default
     */
    public function __construct()
    {
        // Master slide
        $this->masterSlides = array(
            array(
                'masterid' => 1,
                'body' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:bg>
            <p:bgRef idx="1001">
                <a:schemeClr val="bg1" />
            </p:bgRef>
        </p:bg>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title Placeholder 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="274638" />
                        <a:ext cx="8229600" cy="1143000" />
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst />
                    </a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr">
                        <a:normAutofit />
                    </a:bodyPr>
                    <a:lstStyle />
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Text Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="1600200" />
                        <a:ext cx="8229600" cy="4525963" />
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst />
                    </a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0">
                        <a:normAutofit />
                    </a:bodyPr>
                    <a:lstStyle />
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Date Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="2" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="6356350" />
                        <a:ext cx="2133600" cy="365125" />
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst />
                    </a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr" />
                    <a:lstStyle>
                        <a:lvl1pPr algn="l">
                            <a:defRPr sz="1200">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl1pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:fld id="{C6430DBB-9FD5-43E7-88F1-55A569E9525E}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Footer Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="3" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="3124200" y="6356350" />
                        <a:ext cx="2895600" cy="365125" />
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst />
                    </a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr" />
                    <a:lstStyle>
                        <a:lvl1pPr algn="ctr">
                            <a:defRPr sz="1200">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl1pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Slide Number Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="4" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="6553200" y="6356350" />
                        <a:ext cx="2133600" cy="365125" />
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                        <a:avLst />
                    </a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr" />
                    <a:lstStyle>
                        <a:lvl1pPr algn="r">
                            <a:defRPr sz="1200">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl1pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:fld id="{EE336665-E7E9-4861-9ADF-F11A47CBAD79}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>&lt;#&gt;</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink" />
    <p:sldLayoutIdLst>
        <p:sldLayoutId id="2147483649" r:id="rId1" />
        <p:sldLayoutId id="2147483650" r:id="rId2" />
        <p:sldLayoutId id="2147483651" r:id="rId3" />
        <p:sldLayoutId id="2147483652" r:id="rId4" />
        <p:sldLayoutId id="2147483653" r:id="rId5" />
        <p:sldLayoutId id="2147483654" r:id="rId6" />
        <p:sldLayoutId id="2147483655" r:id="rId7" />
        <p:sldLayoutId id="2147483656" r:id="rId8" />
        <p:sldLayoutId id="2147483657" r:id="rId9" />
        <p:sldLayoutId id="2147483658" r:id="rId10" />
        <p:sldLayoutId id="2147483659" r:id="rId11" />
    </p:sldLayoutIdLst>
    <p:txStyles>
        <p:titleStyle>
            <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="0" />
                </a:spcBef>
                <a:buNone />
                <a:defRPr sz="4400" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mj-lt" />
                    <a:ea typeface="+mj-ea" />
                    <a:cs typeface="+mj-cs" />
                </a:defRPr>
            </a:lvl1pPr>
        </p:titleStyle>
        <p:bodyStyle>
            <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="&#149;" />
                <a:defRPr sz="3200" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl1pPr>
            <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="-" />
                <a:defRPr sz="2800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl2pPr>
            <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="&#149;" />
                <a:defRPr sz="2400" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl3pPr>
            <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="-" />
                <a:defRPr sz="2000" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl4pPr>
            <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="&#187;" />
                <a:defRPr sz="2000" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl5pPr>
            <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="&#149;" />
                <a:defRPr sz="2000" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl6pPr>
            <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="&#149;" />
                <a:defRPr sz="2000" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl7pPr>
            <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="&#149;" />
                <a:defRPr sz="2000" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl8pPr>
            <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:spcBef>
                    <a:spcPct val="20000" />
                </a:spcBef>
                <a:buFont typeface="Arial" pitchFamily="34" charset="0" />
                <a:buChar char="&#149;" />
                <a:defRPr sz="2000" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl9pPr>
        </p:bodyStyle>
        <p:otherStyle>
            <a:defPPr>
                <a:defRPr lang="nl-BE" />
            </a:defPPr>
            <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl1pPr>
            <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl2pPr>
            <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl3pPr>
            <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl4pPr>
            <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl5pPr>
            <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl6pPr>
            <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl7pPr>
            <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl8pPr>
            <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
                <a:defRPr sz="1800" kern="1200">
                    <a:solidFill>
                        <a:schemeClr val="tx1" />
                    </a:solidFill>
                    <a:latin typeface="+mn-lt" />
                    <a:ea typeface="+mn-ea" />
                    <a:cs typeface="+mn-cs" />
                </a:defRPr>
            </a:lvl9pPr>
        </p:otherStyle>
    </p:txStyles>
</p:sldMaster>'));

        // Master slide relations
        $this->masterSlideRels = array(
            //array('masterid' => '', 'id' => '', 'type' => '', 'contentType' => '', 'target' => '', 'contents' => '')
        );

        // Theme
        $this->themes = array(
            array(
                'masterid' => 1,
                'body'     => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
 <a:themeElements>
  <a:clrScheme name="Office">
   <a:dk1>
    <a:sysClr val="windowText" lastClr="000000"/>
   </a:dk1>
   <a:lt1>
    <a:sysClr val="window" lastClr="FFFFFF"/>
   </a:lt1>
   <a:dk2>
    <a:srgbClr val="1F497D"/>
   </a:dk2>
   <a:lt2>
    <a:srgbClr val="EEECE1"/>
   </a:lt2>
   <a:accent1>
    <a:srgbClr val="4F81BD"/>
   </a:accent1>
   <a:accent2>
    <a:srgbClr val="C0504D"/>
   </a:accent2>
   <a:accent3>
    <a:srgbClr val="9BBB59"/>
   </a:accent3>
   <a:accent4>
    <a:srgbClr val="8064A2"/>
   </a:accent4>
   <a:accent5>
    <a:srgbClr val="4BACC6"/>
   </a:accent5>
   <a:accent6>
    <a:srgbClr val="F79646"/>
   </a:accent6>
   <a:hlink>
    <a:srgbClr val="0000FF"/>
   </a:hlink>
   <a:folHlink>
    <a:srgbClr val="800080"/>
   </a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="Office">
   <a:majorFont>
    <a:latin typeface="Calibri"/>
    <a:ea typeface=""/>
    <a:cs typeface=""/>
    <a:font script="Jpan" typeface="?? ?????"/>
    <a:font script="Hang" typeface="?? ??"/>
    <a:font script="Hans" typeface="??"/>
    <a:font script="Hant" typeface="????"/>
    <a:font script="Arab" typeface="Times New Roman"/>
    <a:font script="Hebr" typeface="Times New Roman"/>
    <a:font script="Thai" typeface="Angsana New"/>
    <a:font script="Ethi" typeface="Nyala"/>
    <a:font script="Beng" typeface="Vrinda"/>
    <a:font script="Gujr" typeface="Shruti"/>
    <a:font script="Khmr" typeface="MoolBoran"/>
    <a:font script="Knda" typeface="Tunga"/>
    <a:font script="Guru" typeface="Raavi"/>
    <a:font script="Cans" typeface="Euphemia"/>
    <a:font script="Cher" typeface="Plantagenet Cherokee"/>
    <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
    <a:font script="Tibt" typeface="Microsoft Himalaya"/>
    <a:font script="Thaa" typeface="MV Boli"/>
    <a:font script="Deva" typeface="Mangal"/>
    <a:font script="Telu" typeface="Gautami"/>
    <a:font script="Taml" typeface="Latha"/>
    <a:font script="Syrc" typeface="Estrangelo Edessa"/>
    <a:font script="Orya" typeface="Kalinga"/>
    <a:font script="Mlym" typeface="Kartika"/>
    <a:font script="Laoo" typeface="DokChampa"/>
    <a:font script="Sinh" typeface="Iskoola Pota"/>
    <a:font script="Mong" typeface="Mongolian Baiti"/>
    <a:font script="Viet" typeface="Times New Roman"/>
    <a:font script="Uigh" typeface="Microsoft Uighur"/>
   </a:majorFont>
   <a:minorFont>
    <a:latin typeface="Calibri"/>
    <a:ea typeface=""/>
    <a:cs typeface=""/>
    <a:font script="Jpan" typeface="?? ?????"/>
    <a:font script="Hang" typeface="?? ??"/>
    <a:font script="Hans" typeface="??"/>
    <a:font script="Hant" typeface="????"/>
    <a:font script="Arab" typeface="Arial"/>
    <a:font script="Hebr" typeface="Arial"/>
    <a:font script="Thai" typeface="Cordia New"/>
    <a:font script="Ethi" typeface="Nyala"/>
    <a:font script="Beng" typeface="Vrinda"/>
    <a:font script="Gujr" typeface="Shruti"/>
    <a:font script="Khmr" typeface="DaunPenh"/>
    <a:font script="Knda" typeface="Tunga"/>
    <a:font script="Guru" typeface="Raavi"/>
    <a:font script="Cans" typeface="Euphemia"/>
    <a:font script="Cher" typeface="Plantagenet Cherokee"/>
    <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
    <a:font script="Tibt" typeface="Microsoft Himalaya"/>
    <a:font script="Thaa" typeface="MV Boli"/>
    <a:font script="Deva" typeface="Mangal"/>
    <a:font script="Telu" typeface="Gautami"/>
    <a:font script="Taml" typeface="Latha"/>
    <a:font script="Syrc" typeface="Estrangelo Edessa"/>
    <a:font script="Orya" typeface="Kalinga"/>
    <a:font script="Mlym" typeface="Kartika"/>
    <a:font script="Laoo" typeface="DokChampa"/>
    <a:font script="Sinh" typeface="Iskoola Pota"/>
    <a:font script="Mong" typeface="Mongolian Baiti"/>
    <a:font script="Viet" typeface="Arial"/>
    <a:font script="Uigh" typeface="Microsoft Uighur"/>
   </a:minorFont>
  </a:fontScheme>
  <a:fmtScheme name="Office">
   <a:fillStyleLst>
    <a:solidFill>
     <a:schemeClr val="phClr"/>
    </a:solidFill>
    <a:gradFill rotWithShape="1">
     <a:gsLst>
      <a:gs pos="0">
       <a:schemeClr val="phClr">
        <a:tint val="50000"/>
        <a:satMod val="300000"/>
       </a:schemeClr>
      </a:gs>
      <a:gs pos="35000">
       <a:schemeClr val="phClr">
        <a:tint val="37000"/>
        <a:satMod val="300000"/>
       </a:schemeClr>
      </a:gs>
      <a:gs pos="100000">
       <a:schemeClr val="phClr">
        <a:tint val="15000"/>
        <a:satMod val="350000"/>
       </a:schemeClr>
      </a:gs>
     </a:gsLst>
     <a:lin ang="16200000" scaled="1"/>
    </a:gradFill>
    <a:gradFill rotWithShape="1">
     <a:gsLst>
      <a:gs pos="0">
       <a:schemeClr val="phClr">
        <a:shade val="51000"/>
        <a:satMod val="130000"/>
       </a:schemeClr>
      </a:gs>
      <a:gs pos="80000">
       <a:schemeClr val="phClr">
        <a:shade val="93000"/>
        <a:satMod val="130000"/>
       </a:schemeClr>
      </a:gs>
      <a:gs pos="100000">
       <a:schemeClr val="phClr">
        <a:shade val="94000"/>
        <a:satMod val="135000"/>
       </a:schemeClr>
      </a:gs>
     </a:gsLst>
     <a:lin ang="16200000" scaled="0"/>
    </a:gradFill>
   </a:fillStyleLst>
   <a:lnStyleLst>
    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
     <a:solidFill>
      <a:schemeClr val="phClr">
       <a:shade val="95000"/>
       <a:satMod val="105000"/>
      </a:schemeClr>
     </a:solidFill>
     <a:prstDash val="solid"/>
    </a:ln>
    <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
     <a:solidFill>
      <a:schemeClr val="phClr"/>
     </a:solidFill>
     <a:prstDash val="solid"/>
    </a:ln>
    <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
     <a:solidFill>
      <a:schemeClr val="phClr"/>
     </a:solidFill>
     <a:prstDash val="solid"/>
    </a:ln>
   </a:lnStyleLst>
   <a:effectStyleLst>
    <a:effectStyle>
     <a:effectLst>
      <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
       <a:srgbClr val="000000">
        <a:alpha val="38000"/>
       </a:srgbClr>
      </a:outerShdw>
     </a:effectLst>
    </a:effectStyle>
    <a:effectStyle>
     <a:effectLst>
      <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
       <a:srgbClr val="000000">
        <a:alpha val="35000"/>
       </a:srgbClr>
      </a:outerShdw>
     </a:effectLst>
    </a:effectStyle>
    <a:effectStyle>
     <a:effectLst>
      <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
       <a:srgbClr val="000000">
        <a:alpha val="35000"/>
       </a:srgbClr>
      </a:outerShdw>
     </a:effectLst>
     <a:scene3d>
      <a:camera prst="orthographicFront">
       <a:rot lat="0" lon="0" rev="0"/>
      </a:camera>
      <a:lightRig rig="threePt" dir="t">
       <a:rot lat="0" lon="0" rev="1200000"/>
      </a:lightRig>
     </a:scene3d>
     <a:sp3d>
      <a:bevelT w="63500" h="25400"/>
     </a:sp3d>
    </a:effectStyle>
   </a:effectStyleLst>
   <a:bgFillStyleLst>
    <a:solidFill>
     <a:schemeClr val="phClr"/>
    </a:solidFill>
    <a:gradFill rotWithShape="1">
     <a:gsLst>
      <a:gs pos="0">
       <a:schemeClr val="phClr">
        <a:tint val="40000"/>
        <a:satMod val="350000"/>
       </a:schemeClr>
      </a:gs>
      <a:gs pos="40000">
       <a:schemeClr val="phClr">
        <a:tint val="45000"/>
        <a:shade val="99000"/>
        <a:satMod val="350000"/>
       </a:schemeClr>
      </a:gs>
      <a:gs pos="100000">
       <a:schemeClr val="phClr">
        <a:shade val="20000"/>
        <a:satMod val="255000"/>
       </a:schemeClr>
      </a:gs>
     </a:gsLst>
     <a:path path="circle">
      <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
     </a:path>
    </a:gradFill>
    <a:gradFill rotWithShape="1">
     <a:gsLst>
      <a:gs pos="0">
       <a:schemeClr val="phClr">
        <a:tint val="80000"/>
        <a:satMod val="300000"/>
       </a:schemeClr>
      </a:gs>
      <a:gs pos="100000">
       <a:schemeClr val="phClr">
        <a:shade val="30000"/>
        <a:satMod val="200000"/>
       </a:schemeClr>
      </a:gs>
     </a:gsLst>
     <a:path path="circle">
      <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
     </a:path>
    </a:gradFill>
   </a:bgFillStyleLst>
  </a:fmtScheme>
 </a:themeElements>
 <a:objectDefaults/>
 <a:extraClrSchemeLst/>
</a:theme>'));

        // Theme relations
        $this->themeRelations = array(
            //array('masterid' => 1, 'id' => '', 'type' => '', 'contentType' => '', 'target' => '', 'contents' => '')
        );

        // Layouts - Layout::TITLE_SLIDE
        $this->layouts[1] = array(
            'masterid'  => 1,
            'name'      => Layout::TITLE_SLIDE,
            'body'      => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1">
    <p:cSld name="Title Slide">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ctrTitle" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="685800" y="2130425" />
                        <a:ext cx="7772400" cy="1470025" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Subtitle 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="subTitle" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="1371600" y="3886200" />
                        <a:ext cx="6400800" cy="1752600" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr marL="0" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl1pPr>
                        <a:lvl2pPr marL="457200" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl2pPr>
                        <a:lvl3pPr marL="914400" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl3pPr>
                        <a:lvl4pPr marL="1371600" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl4pPr>
                        <a:lvl5pPr marL="1828800" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl5pPr>
                        <a:lvl6pPr marL="2286000" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl6pPr>
                        <a:lvl7pPr marL="2743200" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl7pPr>
                        <a:lvl8pPr marL="3200400" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl8pPr>
                        <a:lvl9pPr marL="3657600" indent="0" algn="ctr">
                            <a:buNone />
                            <a:defRPr>
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master subtitle style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Date Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Footer Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Slide Number Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::TITLE_AND_CONTENT
        $this->layouts[2] = array(
            'masterid'  => 1,
            'name'      => Layout::TITLE_AND_CONTENT,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="obj" preserve="1">
    <p:cSld name="Title and Content">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Content Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Date Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Footer Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Slide Number Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::SECTION_HEADER
        $this->layouts[3] = array(
            'masterid'  => 1,
            'name'      => Layout::SECTION_HEADER,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="secHead" preserve="1">
    <p:cSld name="Section Header">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="722313" y="4406900" />
                        <a:ext cx="7772400" cy="1362075" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr anchor="t" />
                    <a:lstStyle>
                        <a:lvl1pPr algn="l">
                            <a:defRPr sz="4000" b="1" cap="all" />
                        </a:lvl1pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Text Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="722313" y="2906713" />
                        <a:ext cx="7772400" cy="1500187" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr anchor="b" />
                    <a:lstStyle>
                        <a:lvl1pPr marL="0" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl1pPr>
                        <a:lvl2pPr marL="457200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1800">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl2pPr>
                        <a:lvl3pPr marL="914400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl3pPr>
                        <a:lvl4pPr marL="1371600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1400">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl4pPr>
                        <a:lvl5pPr marL="1828800" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1400">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl5pPr>
                        <a:lvl6pPr marL="2286000" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1400">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl6pPr>
                        <a:lvl7pPr marL="2743200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1400">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl7pPr>
                        <a:lvl8pPr marL="3200400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1400">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl8pPr>
                        <a:lvl9pPr marL="3657600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1400">
                                <a:solidFill>
                                    <a:schemeClr val="tx1">
                                        <a:tint val="75000" />
                                    </a:schemeClr>
                                </a:solidFill>
                            </a:defRPr>
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Date Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Footer Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Slide Number Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::TWO_CONTENT
        $this->layouts[4] = array(
            'masterid'  => 1,
            'name'      => Layout::TWO_CONTENT,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="twoObj" preserve="1">
    <p:cSld name="Two Content">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Content Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph sz="half" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="1600200" />
                        <a:ext cx="4038600" cy="4525963" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr>
                            <a:defRPr sz="2800" />
                        </a:lvl1pPr>
                        <a:lvl2pPr>
                            <a:defRPr sz="2400" />
                        </a:lvl2pPr>
                        <a:lvl3pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl3pPr>
                        <a:lvl4pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl4pPr>
                        <a:lvl5pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl5pPr>
                        <a:lvl6pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl6pPr>
                        <a:lvl7pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl7pPr>
                        <a:lvl8pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl8pPr>
                        <a:lvl9pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Content Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph sz="half" idx="2" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="4648200" y="1600200" />
                        <a:ext cx="4038600" cy="4525963" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr>
                            <a:defRPr sz="2800" />
                        </a:lvl1pPr>
                        <a:lvl2pPr>
                            <a:defRPr sz="2400" />
                        </a:lvl2pPr>
                        <a:lvl3pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl3pPr>
                        <a:lvl4pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl4pPr>
                        <a:lvl5pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl5pPr>
                        <a:lvl6pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl6pPr>
                        <a:lvl7pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl7pPr>
                        <a:lvl8pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl8pPr>
                        <a:lvl9pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Date Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Footer Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="7" name="Slide Number Placeholder 6" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::COMPARISON
        $this->layouts[5] = array(
            'masterid'  => 1,
            'name'      => Layout::COMPARISON,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="twoTxTwoObj" preserve="1">
    <p:cSld name="Comparison">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr>
                            <a:defRPr />
                        </a:lvl1pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Text Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="1535113" />
                        <a:ext cx="4040188" cy="639762" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr anchor="b" />
                    <a:lstStyle>
                        <a:lvl1pPr marL="0" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2400" b="1" />
                        </a:lvl1pPr>
                        <a:lvl2pPr marL="457200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000" b="1" />
                        </a:lvl2pPr>
                        <a:lvl3pPr marL="914400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1800" b="1" />
                        </a:lvl3pPr>
                        <a:lvl4pPr marL="1371600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl4pPr>
                        <a:lvl5pPr marL="1828800" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl5pPr>
                        <a:lvl6pPr marL="2286000" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl6pPr>
                        <a:lvl7pPr marL="2743200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl7pPr>
                        <a:lvl8pPr marL="3200400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl8pPr>
                        <a:lvl9pPr marL="3657600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Content Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph sz="half" idx="2" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="2174875" />
                        <a:ext cx="4040188" cy="3951288" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr>
                            <a:defRPr sz="2400" />
                        </a:lvl1pPr>
                        <a:lvl2pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl2pPr>
                        <a:lvl3pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl3pPr>
                        <a:lvl4pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl4pPr>
                        <a:lvl5pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl5pPr>
                        <a:lvl6pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl6pPr>
                        <a:lvl7pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl7pPr>
                        <a:lvl8pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl8pPr>
                        <a:lvl9pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Text Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" sz="quarter" idx="3" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="4645025" y="1535113" />
                        <a:ext cx="4041775" cy="639762" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr anchor="b" />
                    <a:lstStyle>
                        <a:lvl1pPr marL="0" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2400" b="1" />
                        </a:lvl1pPr>
                        <a:lvl2pPr marL="457200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000" b="1" />
                        </a:lvl2pPr>
                        <a:lvl3pPr marL="914400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1800" b="1" />
                        </a:lvl3pPr>
                        <a:lvl4pPr marL="1371600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl4pPr>
                        <a:lvl5pPr marL="1828800" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl5pPr>
                        <a:lvl6pPr marL="2286000" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl6pPr>
                        <a:lvl7pPr marL="2743200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl7pPr>
                        <a:lvl8pPr marL="3200400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl8pPr>
                        <a:lvl9pPr marL="3657600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1600" b="1" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Content Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph sz="quarter" idx="4" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="4645025" y="2174875" />
                        <a:ext cx="4041775" cy="3951288" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr>
                            <a:defRPr sz="2400" />
                        </a:lvl1pPr>
                        <a:lvl2pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl2pPr>
                        <a:lvl3pPr>
                            <a:defRPr sz="1800" />
                        </a:lvl3pPr>
                        <a:lvl4pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl4pPr>
                        <a:lvl5pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl5pPr>
                        <a:lvl6pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl6pPr>
                        <a:lvl7pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl7pPr>
                        <a:lvl8pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl8pPr>
                        <a:lvl9pPr>
                            <a:defRPr sz="1600" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="7" name="Date Placeholder 6" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="8" name="Footer Placeholder 7" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="9" name="Slide Number Placeholder 8" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::TITLE_ONLY
        $this->layouts[6] = array(
            'masterid'  => 1,
            'name'      => Layout::TITLE_ONLY,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="titleOnly" preserve="1">
    <p:cSld name="Title Only">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Date Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Footer Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Slide Number Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::BLANK
        $this->layouts[7] = array(
            'masterid'  => 1,
            'name'      => Layout::BLANK,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
    <p:cSld name="Blank">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Date Placeholder 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Footer Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Slide Number Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::CONTENT_WITH_CAPTION
        $this->layouts[8] = array(
            'masterid'  => 1,
            'name'      => Layout::CONTENT_WITH_CAPTION,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="objTx" preserve="1">
    <p:cSld name="Content with Caption">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="273050" />
                        <a:ext cx="3008313" cy="1162050" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr anchor="b" />
                    <a:lstStyle>
                        <a:lvl1pPr algn="l">
                            <a:defRPr sz="2000" b="1" />
                        </a:lvl1pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Content Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="3575050" y="273050" />
                        <a:ext cx="5111750" cy="5853113" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr>
                            <a:defRPr sz="3200" />
                        </a:lvl1pPr>
                        <a:lvl2pPr>
                            <a:defRPr sz="2800" />
                        </a:lvl2pPr>
                        <a:lvl3pPr>
                            <a:defRPr sz="2400" />
                        </a:lvl3pPr>
                        <a:lvl4pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl4pPr>
                        <a:lvl5pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl5pPr>
                        <a:lvl6pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl6pPr>
                        <a:lvl7pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl7pPr>
                        <a:lvl8pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl8pPr>
                        <a:lvl9pPr>
                            <a:defRPr sz="2000" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Text Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" sz="half" idx="2" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="1435100" />
                        <a:ext cx="3008313" cy="4691063" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr marL="0" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1400" />
                        </a:lvl1pPr>
                        <a:lvl2pPr marL="457200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1200" />
                        </a:lvl2pPr>
                        <a:lvl3pPr marL="914400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1000" />
                        </a:lvl3pPr>
                        <a:lvl4pPr marL="1371600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl4pPr>
                        <a:lvl5pPr marL="1828800" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl5pPr>
                        <a:lvl6pPr marL="2286000" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl6pPr>
                        <a:lvl7pPr marL="2743200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl7pPr>
                        <a:lvl8pPr marL="3200400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl8pPr>
                        <a:lvl9pPr marL="3657600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Date Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Footer Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="7" name="Slide Number Placeholder 6" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::PICTURE_WITH_CAPTION
        $this->layouts[9] = array(
            'masterid'  => 1,
            'name'      => Layout::PICTURE_WITH_CAPTION,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="picTx" preserve="1">
    <p:cSld name="Picture with Caption">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="1792288" y="4800600" />
                        <a:ext cx="5486400" cy="566738" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr anchor="b" />
                    <a:lstStyle>
                        <a:lvl1pPr algn="l">
                            <a:defRPr sz="2000" b="1" />
                        </a:lvl1pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Picture Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="pic" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="1792288" y="612775" />
                        <a:ext cx="5486400" cy="4114800" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr marL="0" indent="0">
                            <a:buNone />
                            <a:defRPr sz="3200" />
                        </a:lvl1pPr>
                        <a:lvl2pPr marL="457200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2800" />
                        </a:lvl2pPr>
                        <a:lvl3pPr marL="914400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2400" />
                        </a:lvl3pPr>
                        <a:lvl4pPr marL="1371600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000" />
                        </a:lvl4pPr>
                        <a:lvl5pPr marL="1828800" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000" />
                        </a:lvl5pPr>
                        <a:lvl6pPr marL="2286000" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000" />
                        </a:lvl6pPr>
                        <a:lvl7pPr marL="2743200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000" />
                        </a:lvl7pPr>
                        <a:lvl8pPr marL="3200400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000" />
                        </a:lvl8pPr>
                        <a:lvl9pPr marL="3657600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="2000" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Text Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" sz="half" idx="2" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="1792288" y="5367338" />
                        <a:ext cx="5486400" cy="804862" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle>
                        <a:lvl1pPr marL="0" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1400" />
                        </a:lvl1pPr>
                        <a:lvl2pPr marL="457200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1200" />
                        </a:lvl2pPr>
                        <a:lvl3pPr marL="914400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="1000" />
                        </a:lvl3pPr>
                        <a:lvl4pPr marL="1371600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl4pPr>
                        <a:lvl5pPr marL="1828800" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl5pPr>
                        <a:lvl6pPr marL="2286000" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl6pPr>
                        <a:lvl7pPr marL="2743200" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl7pPr>
                        <a:lvl8pPr marL="3200400" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl8pPr>
                        <a:lvl9pPr marL="3657600" indent="0">
                            <a:buNone />
                            <a:defRPr sz="900" />
                        </a:lvl9pPr>
                    </a:lstStyle>
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Date Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Footer Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="7" name="Slide Number Placeholder 6" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::TITLE_AND_VERTICAL_TEXT
        $this->layouts[10] = array(
            'masterid'  => 1,
            'name'      => Layout::TITLE_AND_VERTICAL_TEXT,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="vertTx" preserve="1">
    <p:cSld name="Title and Vertical Text">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Vertical Text Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" orient="vert" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr vert="eaVert" />
                    <a:lstStyle />
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Date Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Footer Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Slide Number Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layouts - Layout::VERTICAL_TITLE_AND_TEXT
        $this->layouts[11] = array(
            'masterid'  => 1,
            'name'      => Layout::VERTICAL_TITLE_AND_TEXT,
            'body'      =>  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="vertTitleAndTx" preserve="1">
    <p:cSld name="Vertical Title and Text">
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Vertical Title 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="title" orient="vert" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="6629400" y="274638" />
                        <a:ext cx="2057400" cy="5851525" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="eaVert" />
                    <a:lstStyle />
                    <a:p>
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master title style</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Vertical Text Placeholder 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" orient="vert" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm>
                        <a:off x="457200" y="274638" />
                        <a:ext cx="6019800" cy="5851525" />
                    </a:xfrm>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr vert="eaVert" />
                    <a:lstStyle />
                    <a:p>
                        <a:pPr lvl="0" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Click to edit Master text styles</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="1" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Second level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="2" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Third level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="3" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fourth level</a:t>
                        </a:r>
                    </a:p>
                    <a:p>
                        <a:pPr lvl="4" />
                        <a:r>
                            <a:rPr lang="en-US" smtClean="0" />
                            <a:t>Fifth level</a:t>
                        </a:r>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Date Placeholder 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="dt" sz="half" idx="10" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{6F0DA8BB-0D18-469F-8022-DD923457DE3A}" type="datetimeFigureOut">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>16/04/2009</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Footer Placeholder 4" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="ftr" sz="quarter" idx="11" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="6" name="Slide Number Placeholder 5" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="12" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{B5274F97-0F13-42E5-9A1D-07478243785D}" type="slidenum">
                            <a:rPr lang="nl-BE" smtClean="0" />
                            <a:t>?#?</a:t>
                        </a:fld>
                        <a:endParaRPr lang="nl-BE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:sldLayout>');

        // Layout relations
        $this->layoutRelations = array(
            //array('layoutId' => 0, 'id' => '', 'type' => '', 'contentType' => '', 'target' => '', 'contents' => '')
        );
    }
}
