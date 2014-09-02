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

namespace PhpOffice\PhpPowerpoint\Reader;

use PhpOffice\PhpPowerpoint\Shared\OLERead;
use PhpOffice\PhpPowerpoint\Shape\Drawing;
use PhpOffice\PhpPowerpoint\PhpPowerpoint;
use PhpOffice\PhpPowerpoint\Shape\MemoryDrawing;
use PhpOffice\PhpPowerpoint\Style\Alignment;
use PhpOffice\PhpPowerpoint\Style\Color;
use PhpOffice\PhpPowerpoint\Shape\RichText\Paragraph;
use PhpOffice\PhpPowerpoint\Shape\RichText;
use PhpOffice\PhpPowerpoint\AbstractShape;
use PhpOffice\PhpPowerpoint\Style\Bullet;
use PhpOffice\PhpPowerpoint\Shape\Hyperlink;
use PhpOffice\PhpPowerpoint\Shape\Line;

/**
 * Serialized format reader
 */
class PowerPoint97 implements ReaderInterface
{
    const OFFICEARTBLIPEMF = 0xF01A;
    const OFFICEARTBLIPWMF = 0xF01B;
    const OFFICEARTBLIPPICT = 0xF01C;
    const OFFICEARTBLIPJPG = 0xF01D;
    const OFFICEARTBLIPPNG = 0xF01E;
    const OFFICEARTBLIPDIB = 0xF01F;
    const OFFICEARTBLIPTIFF = 0xF029;
    const OFFICEARTBLIPJPEG = 0xF02A;
    
    /**
     * @link http://msdn.microsoft.com/en-us/library/dd945336(v=office.12).aspx
     */
    const RT_ANIMATIONINFO = 0x1014;
    const RT_ANIMATIONINFOATOM = 0x0FF1;
    const RT_BINARYTAGDATABLOB = 0x138B;
    const RT_BLIPCOLLECTION9 = 0x07F8;
    const RT_BLIPENTITY9ATOM = 0x07F9;
    const RT_BOOKMARKCOLLECTION = 0x07E3;
    const RT_BOOKMARKENTITYATOM = 0x0FD0;
    const RT_BOOKMARKSEEDATOM = 0x07E9;
    const RT_BROADCASTDOCINFO9 = 0x177E;
    const RT_BROADCASTDOCINFO9ATOM = 0x177F;
    const RT_BUILDATOM = 0x2B03;
    const RT_BUILDLIST = 0x2B02;
    const RT_CHARTBUILD = 0x2B04;
    const RT_CHARTBUILDATOM = 0x2B05;
    const RT_COLORSCHEMEATOM = 0x07F0;
    const RT_COMMENT10 = 0x2EE0;
    const RT_COMMENT10ATOM = 0x2EE1;
    const RT_COMMENTINDEX10 = 0x2EE4;
    const RT_COMMENTINDEX10ATOM = 0x2EE5;
    const RT_CRYPTSESSION10CONTAINER = 0x2F14;
    const RT_CURRENTUSERATOM = 0x0FF6;
    const RT_CSTRING = 0x0FBA;
    const RT_DATETIMEMETACHARATOM = 0x0FF7;
    const RT_DEFAULTRULERATOM = 0x0FAB;
    const RT_DOCROUTINGSLIPATOM = 0x0406;
    const RT_DIAGRAMBUILD = 0x2B06;
    const RT_DIAGRAMBUILDATOM = 0x2B07;
    const RT_DIFF10 = 0x2EED;
    const RT_DIFF10ATOM = 0x2EEE;
    const RT_DIFFTREE10 = 0x2EEC;
    const RT_DOCTOOLBARSTATES10ATOM = 0x36B1;
    const RT_DOCUMENT = 0x03E8;
    const RT_DOCUMENTATOM = 0x03E9;
    const RT_DRAWING = 0x040C;
    const RT_DRAWINGGROUP = 0x040B;
    const RT_ENDDOCUMENTATOM = 0x03EA;
    const RT_EXTERNALAVIMOVIE = 0x1006;
    const RT_EXTERNALCDAUDIO = 0x100E;
    const RT_EXTERNALCDAUDIOATOM = 0x1012;
    const RT_EXTERNALHYPERLINK = 0x0FD7;
    const RT_EXTERNALHYPERLINK9 = 0x0FE4;
    const RT_EXTERNALHYPERLINKATOM = 0x0FD3;
    const RT_EXTERNALHYPERLINKFLAGSATOM = 0x1018;
    const RT_EXTERNALMCIMOVIE = 0x1007;
    const RT_EXTERNALMEDIAATOM = 0x1004;
    const RT_EXTERNALMIDIAUDIO = 0x100D;
    const RT_EXTERNALOBJECTLIST = 0x0409;
    const RT_EXTERNALOBJECTLISTATOM = 0x040A;
    const RT_EXTERNALOBJECTREFATOM = 0x0BC1;
    const RT_EXTERNALOLECONTROL = 0x0FEE;
    const RT_EXTERNALOLECONTROLATOM = 0x0FFB;
    const RT_EXTERNALOLEEMBED = 0x0FCC;
    const RT_EXTERNALOLEEMBEDATOM = 0x0FCD;
    const RT_EXTERNALOLELINK = 0x0FCE;
    const RT_EXTERNALOLELINKATOM = 0x0FD1;
    const RT_EXTERNALOLEOBJECTATOM = 0x0FC3;
    const RT_EXTERNALOLEOBJECTSTG = 0x1011;
    const RT_EXTERNALVIDEO = 0x1005;
    const RT_EXTERNALWAVAUDIOEMBEDDED = 0x100F;
    const RT_EXTERNALWAVAUDIOEMBEDDEDATOM = 0x1013;
    const RT_EXTERNALWAVAUDIOLINK = 0x1010;
    const RT_ENVELOPEDATA9ATOM = 0x1785;
    const RT_ENVELOPEFLAGS9ATOM = 0x1784;
    const RT_ENVIRONMENT = 0x03F2;
    const RT_FONTCOLLECTION = 0x07D5;
    const RT_FONTCOLLECTION10 = 0x07D6;
    const RT_FONTEMBEDDATABLOB = 0x0FB8;
    const RT_FONTEMBEDFLAGS10ATOM = 0x32C8;
    const RT_FILTERPRIVACYFLAGS10ATOM = 0x36B0;
    const RT_FONTENTITYATOM = 0x0FB7;
    const RT_FOOTERMETACHARATOM = 0x0FFA;
    const RT_GENERICDATEMETACHARATOM = 0x0FF8;
    const RT_GRIDSPACING10ATOM = 0x040D;
    const RT_GUIDEATOM = 0x03FB;
    const RT_HANDOUT = 0x0FC9;
    const RT_HASHCODEATOM = 0x2B00;
    const RT_HEADERSFOOTERS = 0x0FD9;
    const RT_HEADERSFOOTERSATOM = 0x0FDA;
    const RT_HEADERMETACHARATOM = 0x0FF9;
    const RT_HTMLDOCINFO9ATOM = 0x177B;
    const RT_HTMLPUBLISHINFOATOM = 0x177C;
    const RT_HTMLPUBLISHINFO9 = 0x177D;
    const RT_INTERACTIVEINFO = 0x0FF2;
    const RT_INTERACTIVEINFOATOM = 0x0FF3;
    const RT_KINSOKU = 0x0FC8;
    const RT_KINSOKUATOM = 0x0FD2;
    const RT_LEVELINFOATOM = 0x2B0A;
    const RT_LINKEDSHAPE10ATOM = 0x2EE6;
    const RT_LINKEDSLIDE10ATOM = 0x2EE7;
    const RT_LIST = 0x07D0;
    const RT_MAINMASTER = 0x03F8;
    const RT_MASTERTEXTPROPATOM = 0x0FA2;
    const RT_METAFILE = 0x0FC1;
    const RT_NAMEDSHOW = 0x0411;
    const RT_NAMEDSHOWS = 0x0410;
    const RT_NAMEDSHOWSLIDESATOM = 0x0412;
    const RT_NORMALVIEWSETINFO9 = 0x0414;
    const RT_NORMALVIEWSETINFO9ATOM = 0x0415;
    const RT_NOTES= 0x03F0;
    const RT_NOTESATOM = 0x03F1;
    const RT_NOTESTEXTVIEWINFO9 = 0x0413;
    const RT_OUTLINETEXTPROPS9 = 0x0FAE;
    const RT_OUTLINETEXTPROPS10 = 0x0FB3;
    const RT_OUTLINETEXTPROPS11 = 0x0FB5;
    const RT_OUTLINETEXTPROPSHEADER9ATOM = 0x0FAF;
    const RT_OUTLINETEXTREFATOM = 0x0F9E;
    const RT_OUTLINEVIEWINFO = 0x0407;
    const RT_PERSISTDIRECTORYATOM = 0x1772;
    const RT_PARABUILD = 0x2B08;
    const RT_PARABUILDATOM = 0x2B09;
    const RT_PHOTOALBUMINFO10ATOM = 0x36B2;
    const RT_PLACEHOLDERATOM = 0x0BC3;
    const RT_PRESENTATIONADVISORFLAGS9ATOM = 0x177A;
    const RT_PRINTOPTIONSATOM = 0x1770;
    const RT_PROGBINARYTAG = 0x138A;
    const RT_PROGSTRINGTAG = 0x1389;
    const RT_PROGTAGS = 0x1388;
    const RT_RECOLORINFOATOM = 0x0FE7;
    const RT_RTFDATETIMEMETACHARATOM = 0x1015;
    const RT_ROUNDTRIPANIMATIONATOM12ATOM = 0x2B0B;
    const RT_ROUNDTRIPANIMATIONHASHATOM12ATOM = 0x2B0D;
    const RT_ROUNDTRIPCOLORMAPPING12ATOM = 0x040F;
    const RT_ROUNDTRIPCOMPOSITEMASTERID12ATOM = 0x041D;
    const RT_ROUNDTRIPCONTENTMASTERID12ATOM = 0x0422;
    const RT_ROUNDTRIPCONTENTMASTERINFO12ATOM = 0x041E;
    const RT_ROUNDTRIPCUSTOMTABLESTYLES12ATOM = 0x0428;
    const RT_ROUNDTRIPDOCFLAGS12ATOM = 0x0425;
    const RT_ROUNDTRIPHEADERFOOTERDEFAULTS12ATOM = 0x0424;
    const RT_ROUNDTRIPHFPLACEHOLDER12ATOM = 0x0420;
    const RT_ROUNDTRIPNEWPLACEHOLDERID12ATOM = 0x0BDD;
    const RT_ROUNDTRIPNOTESMASTERTEXTSTYLES12ATOM = 0x0427;
    const RT_ROUNDTRIPOARTTEXTSTYLES12ATOM = 0x0423;
    const RT_ROUNDTRIPORIGINALMAINMASTERID12ATOM = 0x041C;
    const RT_ROUNDTRIPSHAPECHECKSUMFORCL12ATOM = 0x0426;
    const RT_ROUNDTRIPSHAPEID12ATOM = 0x041F;
    const RT_ROUNDTRIPSLIDESYNCINFO12 = 0x3714;
    const RT_ROUNDTRIPSLIDESYNCINFOATOM12 = 0x3715;
    const RT_ROUNDTRIPTHEME12ATOM = 0x040E;
    const RT_SHAPEATOM = 0x0BDB;
    const RT_SHAPEFLAGS10ATOM = 0x0BDC;
    const RT_SLIDE = 0x03EE;
    const RT_SLIDEATOM = 0x03EF;
    const RT_SLIDEFLAGS10ATOM = 0x2EEA;
    const RT_SLIDELISTENTRY10ATOM = 0x2EF0;
    const RT_SLIDELISTTABLE10 = 0x2EF1;
    const RT_SLIDELISTWITHTEXT = 0x0FF0;
    const RT_SLIDELISTTABLESIZE10ATOM = 0x2EEF;
    const RT_SLIDENUMBERMETACHARATOM = 0x0FD8;
    const RT_SLIDEPERSISTATOM = 0x03F3;
    const RT_SLIDESHOWDOCINFOATOM = 0x0401;
    const RT_SLIDESHOWSLIDEINFOATOM = 0x03F9;
    const RT_SLIDETIME10ATOM = 0x2EEB;
    const RT_SLIDEVIEWINFO = 0x03FA;
    const RT_SLIDEVIEWINFOATOM = 0x03FE;
    const RT_SMARTTAGSTORE11CONTAINER = 0x36B3;
    const RT_SOUND = 0x07E6;
    const RT_SOUNDCOLLECTION = 0x07E4;
    const RT_SOUNDCOLLECTIONATOM = 0x07E5;
    const RT_SOUNDDATABLOB = 0x07E7;
    const RT_SORTERVIEWINFO = 0x0408;
    const RT_STYLETEXTPROPATOM = 0x0FA1;
    const RT_STYLETEXTPROP10ATOM = 0x0FB1;
    const RT_STYLETEXTPROP11ATOM = 0x0FB6;
    const RT_STYLETEXTPROP9ATOM = 0x0FAC;
    const RT_SUMMARY = 0x0402;
    const RT_TEXTBOOKMARKATOM = 0x0FA7;
    const RT_TEXTBYTESATOM = 0x0FA8;
    const RT_TEXTCHARFORMATEXCEPTIONATOM = 0x0FA4;
    const RT_TEXTCHARSATOM = 0x0FA0;
    const RT_TEXTDEFAULTS10ATOM = 0x0FB4;
    const RT_TEXTDEFAULTS9ATOM = 0x0FB0;
    const RT_TEXTHEADERATOM = 0x0F9F;
    const RT_TEXTINTERACTIVEINFOATOM = 0x0FDF;
    const RT_TEXTMASTERSTYLEATOM = 0x0FA3;
    const RT_TEXTMASTERSTYLE10ATOM = 0x0FB2;
    const RT_TEXTMASTERSTYLE9ATOM = 0x0FAD;
    const RT_TEXTPARAGRAPHFORMATEXCEPTIONATOM = 0x0FA5;
    const RT_TEXTRULERATOM = 0x0FA6;
    const RT_TEXTSPECIALINFOATOM = 0x0FAA;
    const RT_TEXTSPECIALINFODEFAULTATOM = 0x0FA9;
    const RT_TIMEANIMATEBEHAVIOR = 0xF134;
    const RT_TIMEANIMATEBEHAVIORCONTAINER = 0xF12B;
    const RT_TIMEANIMATIONVALUE = 0xF143;
    const RT_TIMEANIMATIONVALUELIST = 0xF13F;
    const RT_TIMEBEHAVIOR = 0xF133;
    const RT_TIMEBEHAVIORCONTAINER = 0xF12A;
    const RT_TIMECOLORBEHAVIOR = 0xF135;
    const RT_TIMECOLORBEHAVIORCONTAINER = 0xF12C;
    const RT_TIMECLIENTVISUALELEMENT = 0xF13C;
    const RT_TIMECOMMANDBEHAVIOR = 0xF13B;
    const RT_TIMECOMMANDBEHAVIORCONTAINER = 0xF132;
    const RT_TIMECONDITION = 0xF128;
    const RT_TIMECONDITIONCONTAINER = 0xF125;
    const RT_TIMEEFFECTBEHAVIOR = 0xF136;
    const RT_TIMEEFFECTBEHAVIORCONTAINER = 0xF12D;
    const RT_TIMEEXTTIMENODECONTAINER = 0xF144;
    const RT_TIMEITERATEDATA = 0xF140;
    const RT_TIMEMODIFIER = 0xF129;
    const RT_TIMEMOTIONBEHAVIOR = 0xF137;
    const RT_TIMEMOTIONBEHAVIORCONTAINER = 0xF12E;
    const RT_TIMENODE = 0xF127;
    const RT_TIMEPROPERTYLIST = 0xF13D;
    const RT_TIMEROTATIONBEHAVIOR = 0xF138;
    const RT_TIMEROTATIONBEHAVIORCONTAINER = 0xF12F;
    const RT_TIMESCALEBEHAVIOR = 0xF139;
    const RT_TIMESCALEBEHAVIORCONTAINER = 0xF130;
    const RT_TIMESEQUENCEDATA = 0xF141;
    const RT_TIMESETBEHAVIOR = 0xF13A;
    const RT_TIMESETBEHAVIORCONTAINER = 0xF131;
    const RT_TIMESUBEFFECTCONTAINER = 0xF145;
    const RT_TIMEVARIANT = 0xF142;
    const RT_TIMEVARIANTLIST = 0xF13E;
    const RT_USEREDITATOM = 0x0FF5;
    const RT_VBAINFO = 0x03FF;
    const RT_VBAINFOATOM = 0x0400;
    const RT_VIEWINFOATOM = 0x03FD;
    const RT_VISUALPAGEATOM = 0x2B01;
    const RT_VISUALSHAPEATOM = 0x2AFB;
    
    /**
     * @var http://msdn.microsoft.com/en-us/library/dd926394(v=office.12).aspx
     */
    const SL_BIGOBJECT = 0x0000000F;
    const SL_BLANK = 0x00000010;
    const SL_COLUMNTWOROWS = 0x0000000A;
    const SL_FOUROBJECTS = 0x0000000E;
    const SL_MASTERTITLE = 0x00000002;
    const SL_TITLEBODY = 0x00000001;
    const SL_TITLEONLY = 0x00000007;
    const SL_TITLESLIDE = 0x00000000;
    const SL_TWOCOLUMNS = 0x00000008;
    const SL_TWOCOLUMNSROW = 0x0000000D;
    const SL_TWOROWS = 0x00000009;
    const SL_TWOROWSCOLUMN = 0x0000000B;
    const SL_VERTICALTITLEBODY = 0x00000011;
    const SL_VERTICALTWOROWS = 0x00000012;
    
    /**
     * Array with Fonts
     */
    private $arrayFonts = array();
    /**
     * Array with Hyperlinks
     */
    private $arrayHyperlinks = array();
    /**
     * Array with Pictures
     */
    private $arrayPictures = array();
    /**
     * Offset (in bytes) from the beginning of the PowerPoint Document Stream to the UserEditAtom record for the most recent user edit.
     * @var int
     */
    private $offsetToCurrentEdit;
    /**
     * A structure that specifies a compressed table of sequential persist object identifiers and stream offsets to associated persist objects.
     * @var int[]
     */
    private $rgPersistDirEntry;
    /**
     * Offset (in bytes) from the beginning of the PowerPoint Document Stream to the PersistDirectoryAtom record for this user edit
     * @var int[]
     */
    private $offsetPersistDirectory;
    /**
     * Output Object
     * @var PhpPowerpoint
     */
    private $oPhpPowerpoint;
    /**
     * Stream "Powerpoint Document"
     * @var string
     */
    private $streamPowerpointDocument;
    /**
     * Stream "Current User"
     * @var string
     */
    private $streamCurrentUser;
    /**
     * Stream "Summary Information"
     * @var string
     */
    private $streamSummaryInformation;
    /**
     * Stream "Document Summary Information"
     * @var string
     */
    private $streamDocumentSummaryInformation;
    /**
     * Stream "Pictures"
     * @var string
     */
    private $streamPictures;
    
    /**
     * Can the current \PhpOffice\PhpPowerpoint\Reader\ReaderInterface read the file?
     *
     * @param  string $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function canRead($pFilename)
    {
        return $this->fileSupportsUnserializePHPPowerPoint($pFilename);
    }

    /**
     * Does a file support UnserializePHPPowerPoint ?
     *
     * @param  string    $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function fileSupportsUnserializePHPPowerPoint($pFilename = '')
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        try {
            // Use ParseXL for the hard work.
            $ole = new OLERead();
            // get excel data
            $ole->read($pFilename);
            return true;
        } catch (\Exception $e) {
            return false;
        }
    }

    /**
     * Loads PHPPowerPoint Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPowerpoint\PhpPowerpoint
     * @throws \Exception
     */
    public function load($pFilename)
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new \Exception("Could not open " . $pFilename . " for reading! File does not exist.");
        }

        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePHPPowerPoint($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPowerpoint\Reader\Serialized: " . $pFilename . ".");
        }

        return $this->loadFile($pFilename);
    }

    /**
     * Load PHPPowerPoint Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPowerpoint\PhpPowerpoint
     */
    private function loadFile($pFilename)
    {
        $this->oPhpPowerpoint = new PhpPowerpoint();
        $this->oPhpPowerpoint->removeSlideByIndex();
        
        // Read OLE Blocks
        $this->loadOLE($pFilename);
        // Read pictures in the Pictures Stream
        $this->loadPicturesStream();
        // Read information in the Current User Stream
        $this->loadCurrentUserStream();
        // Read information in the PowerPoint Document Stream
        $this->loadPowerpointDocumentStream();
        
        return $this->oPhpPowerpoint;
    }
    
    /**
     * Read OLE Part
     * @param unknown $pFilename
     */
    private function loadOLE($pFilename)
    {
        // OLE reader
        $oOLE = new OLERead();
        $oOLE->read($pFilename);

        // PowerPoint Document Stream
        $this->streamPowerpointDocument = $oOLE->getStream($oOLE->powerpointDocument);
        
        // Current User Stream
        $this->streamCurrentUser = $oOLE->getStream($oOLE->currentUser);
        
        // Get summary information data
        $this->streamSummaryInformation = $oOLE->getStream($oOLE->summaryInformation);
        
        // Get additional document summary information data
        $this->streamDocumentSummaryInformation = $oOLE->getStream($oOLE->docSummaryInfos);
        
        // Get pictures data
        $this->streamPictures = $oOLE->getStream($oOLE->pictures);
    }

    /**
     * Stream Pictures
     * @link http://msdn.microsoft.com/en-us/library/dd920746(v=office.12).aspx
     */
    private function loadPicturesStream()
    {
        $pos = 0;
        $readSuccess = true;
        do {
            $arrayRH = $this->loadRecordHeader($this->streamPictures, $pos);
            $pos += 8;
            if ($arrayRH['recVer'] == 0x00 && ($arrayRH['recType'] == 0xF007 || ($arrayRH['recType'] >= 0xF018 && $arrayRH['recType'] <= 0xF117))) {
                //@link : http://msdn.microsoft.com/en-us/library/dd950560(v=office.12).aspx
                if ($arrayRH['recType'] == 0xF007) {
                    // OfficeArtFBSE
                    throw new \Exception('Feature not implemented (l.'.__LINE__.')');
                }
                if ($arrayRH['recType'] >= 0xF018 && $arrayRH['recType'] <= 0xF117) {
                    //@link : http://msdn.microsoft.com/en-us/library/dd910081(v=office.12).aspx
                    switch ($arrayRH['recType']) {
                        case self::OFFICEARTBLIPJPG:
                        case self::OFFICEARTBLIPPNG:
                            // rgbUid1
                            $pos += 16;
                            $arrayRH['recLen'] -= 16;
                            if ($arrayRH['recInstance'] == 0x6E1) {
                                // rgbUid2
                                $pos += 16;
                                $arrayRH['recLen'] -= 16;
                            }
                            // tag
                            $pos += 1;
                            $arrayRH['recLen'] -= 1;
                            // BLIPFileData
                            $this->arrayPictures[] = substr($this->streamPictures, $pos, $arrayRH['recLen']);
                            $pos += $arrayRH['recLen'];
                            break;
                        default:
                            throw new \Exception('Feature not implemented (l.'.__LINE__.' : '.dechex($arrayRH['recType'].')'));
                            break;
                    }
                }
            } else {
                $readSuccess = false;
            }
        } while ($readSuccess == true);
    }
    
    /**
     * Stream Current User
     * @link http://msdn.microsoft.com/en-us/library/dd908567(v=office.12).aspx
     */
    private function loadCurrentUserStream()
    {
        $pos = 0;
        
        /**
         * CurrentUserAtom : http://msdn.microsoft.com/en-us/library/dd948895(v=office.12).aspx
         */
        // RecordHeader : http://msdn.microsoft.com/en-us/library/dd926377(v=office.12).aspx
        $rHeader = $this->loadRecordHeader($this->streamCurrentUser, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0x0 || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_CURRENTUSERATOM) {
            throw new \Exception('File PowerPoint 97 in error (Location : CurrentUserAtom > RecordHeader).');
        }

        // Size
        $size = self::getInt4d($this->streamCurrentUser, $pos);
        $pos += 4;
        if ($size !=  0x00000014) {
            throw new \Exception('File PowerPoint 97 in error (Location : CurrentUserAtom > Size).');
        }

        // headerToken
        $headerToken = self::getInt4d($this->streamCurrentUser, $pos);
        $pos += 4;
        if ($headerToken == 0xF3D1C4DF && $headerToken != 0xE391C05F) {
            throw new \Exception('Feature not implemented (l.'.__LINE__.') : Encrypted file');
        }

        // offsetToCurrentEdit
        $this->offsetToCurrentEdit = self::getInt4d($this->streamCurrentUser, $pos);
        $pos += 4;

        // lenUserName
        $lenUserName = self::getInt2d($this->streamCurrentUser, $pos);
        $pos += 2;
        if ($lenUserName > 255) {
            throw new \Exception('File PowerPoint 97 in error (Location : CurrentUserAtom > lenUserName).');
        }

        // docFileVersion
        $docFileVersion = self::getInt2d($this->streamCurrentUser, $pos);
        $pos += 2;
        if ($docFileVersion != 0x03F4) {
            throw new \Exception('File PowerPoint 97 in error (Location : CurrentUserAtom > docFileVersion).');
        }

        // majorVersion
        $majorVersion = self::getInt1d($this->streamCurrentUser, $pos);
        $pos += 1;
        if ($majorVersion != 0x03) {
            throw new \Exception('File PowerPoint 97 in error (Location : CurrentUserAtom > majorVersion).');
        }

        // minorVersion
        $minorVersion = self::getInt1d($this->streamCurrentUser, $pos);
        $pos += 1;
        if ($minorVersion != 0x00) {
            throw new \Exception('File PowerPoint 97 in error (Location : CurrentUserAtom > minorVersion).');
        }
        
        // unused
        $pos += 2;
        
        // ansiUserName
        $ansiUserName = '';
        $char = false;
        do {
            $char = self::getInt1d($this->streamCurrentUser, $pos);
            if (($char >= 0x00 && $char <= 0x1F) || ($char >= 0x7F && $char <= 0x9F)) {
                $char = false;
            } else {
                $ansiUserName .= chr($char);
                $pos += 1;
            }
        } while ($char != false);

        // relVersion
        $relVersion = self::getInt4d($this->streamCurrentUser, $pos);
        $pos += 4;
        if ($relVersion != 0x00000008 && $relVersion != 0x00000009) {
            throw new \Exception('File PowerPoint 97 in error (Location : CurrentUserAtom > relVersion).');
        }
        
        // unicodeUserName
        $unicodeUserName = '';
        for ($inc = 0; $inc < $lenUserName; $inc++) {
            $char = self::getInt2d($this->streamCurrentUser, $pos);
            if (($char >= 0x00 && $char <= 0x1F) || ($char >= 0x7F && $char <= 0x9F)) {
                break;
            }
            $unicodeUserName .= chr($char);
            $pos += 2;
        }
    }
    
    /**
     * Stream Powerpoint Document
     * @link http://msdn.microsoft.com/en-us/library/dd921564(v=office.12).aspx
     */
    private function loadPowerpointDocumentStream()
    {
        $this->loadUserEditAtom();
        
        $this->loadPersistDirectoryAtom();
        
        foreach ($this->rgPersistDirEntry as $offsetDir) {
            $pos = $offsetDir;
            
            $rh = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
            $pos += 8;
            switch ($rh['recType']) {
                case self::RT_DOCUMENT:
                    $this->readRTDocument($pos);
                    break;
                case self::RT_SLIDE:
                    $this->readRTSlide($pos);
                    break;
                case self::RT_MAINMASTER:
                case self::RT_NOTES:
                default:
                    break;
            }
        }
    }
    
    /**
     * UserEditAtom
     * @link http://msdn.microsoft.com/en-us/library/dd945746(v=office.12).aspx
     * @throws \Exception
     */
    private function loadUserEditAtom()
    {
        $pos = $this->offsetToCurrentEdit;
        
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0x0 || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_USEREDITATOM || ($rHeader['recLen'] != 0x0000001C && $rHeader['recLen'] != 0x00000020)) {
            throw new \Exception('File PowerPoint 97 in error (Location : UserEditAtom > RecordHeader).');
        }
        
        // lastSlideIdRef
        $pos += 4;
        // version
        $pos += 2;
        
        // minorVersion
        $minorVersion = self::getInt1d($this->streamPowerpointDocument, $pos);
        $pos += 1;
        if ($minorVersion != 0x00) {
            throw new \Exception('File PowerPoint 97 in error (Location : UserEditAtom > minorVersion).');
        }
        
        // majorVersion
        $majorVersion = self::getInt1d($this->streamPowerpointDocument, $pos);
        $pos += 1;
        if ($majorVersion != 0x03) {
            throw new \Exception('File PowerPoint 97 in error (Location : UserEditAtom > majorVersion).');
        }
        
        // offsetLastEdit
        $pos += 4;
        // offsetPersistDirectory
        $this->offsetPersistDirectory  = self::getInt4d($this->streamPowerpointDocument, $pos);
        $pos += 4;
        
        // docPersistIdRef
        $docPersistIdRef  = self::getInt4d($this->streamPowerpointDocument, $pos);
        $pos += 4;
        if ($docPersistIdRef != 0x00000001) {
            throw new \Exception('File PowerPoint 97 in error (Location : UserEditAtom > docPersistIdRef).');
        }
        
        // persistIdSeed
        $pos += 4;
        // lastView
        $pos += 2;
        // unused
        $pos += 2;
    }
    
    /**
     * PersistDirectoryAtom 
     * @link http://msdn.microsoft.com/en-us/library/dd952680(v=office.12).aspx
     * @throws \Exception
     */
    private function loadPersistDirectoryAtom()
    {
        $pos = $this->offsetPersistDirectory;
        
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0x0 || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_PERSISTDIRECTORYATOM) {
            throw new \Exception('File PowerPoint 97 in error (Location : PersistDirectoryAtom > RecordHeader).');
        }
        // rgPersistDirEntry
        // @link : http://msdn.microsoft.com/en-us/library/dd947347(v=office.12).aspx
        do {
            $data = self::getInt4d($this->streamPowerpointDocument, $pos);
            $pos += 4;
            $rHeader['recLen'] -= 4;
            //$persistId  = ($data >> 0) & bindec('11111111111111111111');
            $cPersist  = ($data >> 20) & bindec('111111111111');
        
            $rgPersistOffset = array();
            for ($inc = 0; $inc < $cPersist; $inc++) {
                $rgPersistOffset[] = self::getInt4d($this->streamPowerpointDocument, $pos);
                $pos += 4;
                $rHeader['recLen'] -= 4;
            }
        } while ($rHeader['recLen'] > 0);
        $this->rgPersistDirEntry = $rgPersistOffset;
    }
    
    /**
     * SlideContainer
     * @link http://msdn.microsoft.com/en-us/library/dd946323(v=office.12).aspx
     * @param int $pos
     */
    private function readRTSlide($pos)
    {
        $oSlide = $this->oPhpPowerpoint->createSlide();
        // echo '@slide'.EOL;
        
        // *** slideAtom (32 bytes)
        // slideAtom > rh
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0x2 || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_SLIDEATOM) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > RecordHeader).');
        }
        
        // slideAtom > geom
        $pos += 4;
        
        // slideAtom > rgPlaceholderTypes
        $rgPlaceholderTypes = array();
        for ($inc = 0; $inc < 8; $inc++) {
            $rgPlaceholderTypes[] = self::getInt1d($this->streamPowerpointDocument, $pos);
            $pos += 1;
        }
        
        // slideAtom > masterIdRef
        $pos += 4;
        // slideAtom > notesIdRef
        $pos += 4;
        // slideAtom > slideFlags
        $pos += 2;
        // slideAtom > unused;
        $pos += 2;
        // *** slideShowSlideInfoAtom (24 bytes)
        $pos += 24;
        
        // perSlideHFContainer (variable)
        // perSlideHFContainer > rh
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0xF || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_HEADERSFOOTERS) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > perSlideHFContainer > RT_HeadersFooters).');
        }
        $pos += $rHeader['recLen'];
        
        // *** rtSlideSyncInfo12 (variable)
        // *** drawing (variable)
        // drawing > rh
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0xF || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_DRAWING) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > drawing > rh).');
        }
        // print_r('@PPDrawing'.EOL);

        // drawing > OfficeArtDg
        // drawing > OfficeArtDg > rh
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0xF || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != 0xF002) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > drawing > OfficeArtDg > rh).');
        }
        
        // drawing > OfficeArtDg > drawingData
        // drawing > OfficeArtDg > drawingData > rh
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0x0 || $rHeader['recInstance'] >= 0xFFE || $rHeader['recType'] != 0xF008 || $rHeader['recLen'] != 0x00000008) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > drawing > OfficeArtDg > drawingData > rh).');
        }
        
        // drawing > OfficeArtDg > drawingData > csp
        $pos += 4;
        // drawing > OfficeArtDg > drawingData > spidCur
        $pos += 4;
        
        // drawing > OfficeArtDg > groupShape
        // drawing > OfficeArtDg > groupShape > rh
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0xF || $rHeader['recInstance'] >= 0xF000 || $rHeader['recType'] != 0xF003) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > drawing > OfficeArtDg > groupShape > rh).');
        }
        
        // drawing > OfficeArtDg > groupShape > rgfb
        do {
            $shape = null;

            $arrShpPrimaryOpt = array();
            $arrClientAnchor = array();
            $arrClientTextBox = array();
            $arrClientTextBox['hyperlink'] = array();
            $arrClientTextBox['text'] = '';
            $arrClientTextBox['numParts'] = 0;
            $arrClientTextBox['numTexts'] = 0;
            
            $rhFB = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
            $pos += 8;
            $rHeader['recLen'] -= 8;

            // print_r(EOL);
            // print_r($rhFB);
            // print_r(EOL);
            if ($rhFB['recVer'] != 0xF || $rhFB['recInstance'] != 0x0000 || ($rhFB['recType'] != 0xF003 && $rhFB['recType'] != 0xF004)) {
                throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > drawing > OfficeArtDg > groupShape > rgfb).');
            }
            
            switch ($rhFB['recType']) {
                case 0xF003: // OfficeArtSpgrContainer
                    // OfficeArtSpgrContainer
                    // print_r('@OfficeArtSpgrContainer'.EOL);
                    break;
                case 0xF004: // OfficeArtSpContainer
                    // shapeGroup
                    $shapeGroup = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    if ($shapeGroup['recVer'] == 0x1 && $shapeGroup['recInstance'] == 0x0000 && $shapeGroup['recType'] == 0xF009 && $shapeGroup['recLen'] == 0x00000010) {
                        // print_r('$shapeGroup'.EOL);
                        $pos += 8;
                        $rHeader['recLen'] -= 8;
                        $arrShapeGroup['xLeft'] = self::getInt4d($this->streamPowerpointDocument, $pos);
                        $pos += 4;
                        $rHeader['recLen'] -= 4;
                        $arrShapeGroup['yTop'] = self::getInt4d($this->streamPowerpointDocument, $pos);
                        $pos += 4;
                        $rHeader['recLen'] -= 4;
                        $arrShapeGroup['xRight'] = self::getInt4d($this->streamPowerpointDocument, $pos);
                        $pos += 4;
                        $rHeader['recLen'] -= 4;
                        $arrShapeGroup['yBottom'] = self::getInt4d($this->streamPowerpointDocument, $pos);
                        $pos += 4;
                        $rHeader['recLen'] -= 4;
                    }
                    
                    // shapeProp
                    $shapeProp = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    if ($shapeProp['recVer'] == 0x2 && $shapeProp['recType'] == 0xF00A && $shapeProp['recLen'] == 0x00000008) {
                        $pos += 8;
                        $rHeader['recLen'] -= 8;
                        // print_r('$shapeProp'.EOL);
                        
                        // spid
                        $pos += 4;
                        $rHeader['recLen'] -= 4;
                        // data
                        $pos += 4;
                        $rHeader['recLen'] -= 4;
                    }
                    
                    // shapePrimaryOptions
                    $shapePrimaryOptions = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    if ($shapePrimaryOptions['recVer'] == 0x3 && $shapePrimaryOptions['recType'] == 0xF00B) {
                        $pos += 8;
                        $rHeader['recLen'] -= 8;
                        // print_r('$shapePrimaryOptions'.EOL);
                        //@link : http://msdn.microsoft.com/en-us/library/dd906086(v=office.12).aspx
                        $officeArtFOPTE = array();
                        for ($inc = 0; $inc < $shapePrimaryOptions['recInstance']; $inc++) {
                            $opid = self::getInt2d($this->streamPowerpointDocument, $pos);
                            $pos += 2;
                            $rHeader['recLen'] -= 2;
                            $shapePrimaryOptions['recLen'] -= 2;
                            $optOp = self::getInt4d($this->streamPowerpointDocument, $pos);
                            $pos += 4;
                            $rHeader['recLen'] -= 4;
                            $shapePrimaryOptions['recLen'] -= 4;
                            $officeArtFOPTE[] = array(
                                'opid' => ($opid >> 0) & bindec('11111111111111'),
                                'fBid' => ($opid >> 14) & bindec('1'),
                                'fComplex' => ($opid >> 15) & bindec('1'),
                                'op' => $optOp,
                            );
                        }
                        //@link : http://code.metager.de/source/xref/kde/calligra/filters/libmso/OPID
                        foreach ($officeArtFOPTE as $opt) {
                            switch ($opt['opid']) {
                                case 0x007F:
                                    // Transform : Protection Boolean Properties
                                    //@link : http://msdn.microsoft.com/en-us/library/dd909131(v=office.12).aspx
                                    break;
                                case 0x0080:
                                    // Text : ltxid
                                    //@link : http://msdn.microsoft.com/en-us/library/dd947446(v=office.12).aspx
                                    break;
                                case 0x0081:
                                    // Text : dxTextLeft
                                    //@link : http://msdn.microsoft.com/en-us/library/dd953234(v=office.12).aspx
                                    $arrShpPrimaryOpt['insetLeft'] = \PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']);
                                    break;
                                case 0x0082:
                                    // Text : dyTextTop
                                    //@link : http://msdn.microsoft.com/en-us/library/dd925068(v=office.12).aspx
                                    $arrShpPrimaryOpt['insetTop'] = \PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']);
                                    break;
                                case 0x0083:
                                    // Text : dxTextRight
                                    //@link : http://msdn.microsoft.com/en-us/library/dd906782(v=office.12).aspx
                                    $arrShpPrimaryOpt['insetRight'] = \PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']);
                                    break;
                                case 0x0084:
                                    // Text : dyTextBottom
                                    //@link : http://msdn.microsoft.com/en-us/library/dd772858(v=office.12).aspx
                                    $arrShpPrimaryOpt['insetBottom'] = \PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']);
                                    break;
                                case 0x0085:
                                    // Text : WrapText
                                    //@link : http://msdn.microsoft.com/en-us/library/dd924770(v=office.12).aspx
                                    break;
                                case 0x0087:
                                    // Text : anchorText
                                    //@link : http://msdn.microsoft.com/en-us/library/dd948575(v=office.12).aspx
                                    break;
                                case 0x00BF:
                                    // Text : Text Boolean Properties
                                    //@link : http://msdn.microsoft.com/en-us/library/dd950905(v=office.12).aspx
                                    break;
                                case 0x0104:
                                    // Blip : pib
                                    //@link : http://msdn.microsoft.com/en-us/library/dd772837(v=office.12).aspx
                                    if ($opt['fComplex'] == 0) {
                                        $arrShpPrimaryOpt['pib'] = $opt['op'];
                                    } else {
                                       // pib Complex
                                    }
                                    break;
                                case 0x140:
                                    // Geometry : geoLeft
                                    //@link : http://msdn.microsoft.com/en-us/library/dd947489(v=office.12).aspx
                                    // print_r('geoLeft : '.$opt['op'].EOL);
                                    break;
                                case 0x141:
                                    // Geometry : geoTop
                                    //@link : http://msdn.microsoft.com/en-us/library/dd949459(v=office.12).aspx
                                    // print_r('geoTop : '.$opt['op'].EOL);
                                    break;
                                case 0x142:
                                    // Geometry : geoRight
                                    //@link : http://msdn.microsoft.com/en-us/library/dd947117(v=office.12).aspx
                                    // print_r('geoRight : '.$opt['op'].EOL);
                                    break;
                                case 0x143:
                                    // Geometry : geoBottom
                                    //@link : http://msdn.microsoft.com/en-us/library/dd948602(v=office.12).aspx
                                    // print_r('geoBottom : '.$opt['op'].EOL);
                                    break;
                                case 0x144:
                                    // Geometry : shapePath
                                    //@link : http://msdn.microsoft.com/en-us/library/dd945249(v=office.12).aspx
                                    $arrShpPrimaryOpt['line'] = true;
                                    break;
                                case 0x145:
                                    // Geometry : pVertices
                                    //@link : http://msdn.microsoft.com/en-us/library/dd949814(v=office.12).aspx
                                    if ($opt['fComplex'] == 1) {
                                        $pos += $opt['op'];
                                        $rHeader['recLen'] -= $opt['op'];
                                        $shapePrimaryOptions['recLen'] -= $opt['op'];
                                    }
                                    break;
                                case 0x146:
                                    // Geometry : pSegmentInfo
                                    //@link : http://msdn.microsoft.com/en-us/library/dd905742(v=office.12).aspx
                                    if ($opt['fComplex'] == 1) {
                                        $pos += $opt['op'];
                                        $rHeader['recLen'] -= $opt['op'];
                                        $shapePrimaryOptions['recLen'] -= $opt['op'];
                                    }
                                    break;
                                case 0x155:
                                    // Geometry : pAdjustHandles
                                    //@link : http://msdn.microsoft.com/en-us/library/dd905890(v=office.12).aspx
                                    if ($opt['fComplex'] == 1) {
                                        $pos += $opt['op'];
                                        $rHeader['recLen'] -= $opt['op'];
                                        $shapePrimaryOptions['recLen'] -= $opt['op'];
                                    }
                                    break;
                                case 0x156:
                                    // Geometry : pGuides
                                    //@link : http://msdn.microsoft.com/en-us/library/dd910801(v=office.12).aspx
                                    if ($opt['fComplex'] == 1) {
                                        $pos += $opt['op'];
                                        $rHeader['recLen'] -= $opt['op'];
                                        $shapePrimaryOptions['recLen'] -= $opt['op'];
                                    }
                                    break;
                                case 0x157:
                                    // Geometry : pInscribe
                                    //@link : http://msdn.microsoft.com/en-us/library/dd904889(v=office.12).aspx
                                    if ($opt['fComplex'] == 1) {
                                        $pos += $opt['op'];
                                        $rHeader['recLen'] -= $opt['op'];
                                        $shapePrimaryOptions['recLen'] -= $opt['op'];
                                    }
                                    break;
                                case 0x17F:
                                    // Geometry Boolean Properties
                                    //@link : http://msdn.microsoft.com/en-us/library/dd944968(v=office.12).aspx
                                    break;
                                case 0x0180:
                                    // Fill : fillType
                                    //@link : http://msdn.microsoft.com/en-us/library/dd947909(v=office.12).aspx
                                    break;
                                case 0x0181:
                                    // Fill : fillColor
                                    //@link : http://msdn.microsoft.com/en-us/library/dd921332(v=office.12).aspx
                                    $red = ($opt['op'] >> 0) & bindec('11111111');
                                    $green = ($opt['op'] >> 8) & bindec('11111111');
                                    $blue = ($opt['op'] >> 16) & bindec('11111111');
                                    
                                    $strColor = str_pad(dechex($red), 2, STR_PAD_LEFT, '0');
                                    $strColor .= str_pad(dechex($green), 2, STR_PAD_LEFT, '0');
                                    $strColor .= str_pad(dechex($blue), 2, STR_PAD_LEFT, '0');
                                    // echo 'fillColor  : '.$strColor.EOL;
                                    break;
                                case 0x0183:
                                    // Fill : fillBackColor
                                    //@link : http://msdn.microsoft.com/en-us/library/dd950634(v=office.12).aspx
                                    $red = ($opt['op'] >> 0) & bindec('11111111');
                                    $green = ($opt['op'] >> 8) & bindec('11111111');
                                    $blue = ($opt['op'] >> 16) & bindec('11111111');
                                    
                                    $strColor = str_pad(dechex($red), 2, STR_PAD_LEFT, '0');
                                    $strColor .= str_pad(dechex($green), 2, STR_PAD_LEFT, '0');
                                    $strColor .= str_pad(dechex($blue), 2, STR_PAD_LEFT, '0');
                                    // echo 'fillBackColor  : '.$strColor.EOL;
                                    break;
                                case 0x0193:
                                    // Fill : fillRectRight
                                    //@link : http://msdn.microsoft.com/en-us/library/dd951294(v=office.12).aspx
                                    // echo 'fillRectRight  : '.\PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']).EOL;
                                    break;
                                case 0x0194:
                                    // Fill : fillRectBottom
                                    //@link : http://msdn.microsoft.com/en-us/library/dd910194(v=office.12).aspx
                                    // echo 'fillRectBottom   : '.\PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']).EOL;
                                    break;
                                case 0x01BF:
                                    // Fill : Fill Style Boolean Properties
                                    //@link : http://msdn.microsoft.com/en-us/library/dd909380(v=office.12).aspx
                                    break;
                                case 0x01C0:
                                    // Line Style : lineColor
                                    //@link : http://msdn.microsoft.com/en-us/library/dd920397(v=office.12).aspx
                                    $red = ($opt['op'] >> 0) & bindec('11111111');
                                    $green = ($opt['op'] >> 8) & bindec('11111111');
                                    $blue = ($opt['op'] >> 16) & bindec('11111111');
                                    
                                    $strColor = str_pad(dechex($red), 2, STR_PAD_LEFT, '0');
                                    $strColor .= str_pad(dechex($green), 2, STR_PAD_LEFT, '0');
                                    $strColor .= str_pad(dechex($blue), 2, STR_PAD_LEFT, '0');
                                    
                                    $arrShpPrimaryOpt['lineColor'] = $strColor;
                                    break;
                                case 0x01C1:
                                    // Line Style : lineOpacity
                                    //@link : http://msdn.microsoft.com/en-us/library/dd923433(v=office.12).aspx
                                    // echo 'lineOpacity : '.dechex($opt['op']).EOL;
                                    break;
                                case 0x01C2:
                                    // Line Style : lineBackColor
                                    //@link : http://msdn.microsoft.com/en-us/library/dd947669(v=office.12).aspx
                                    break;
                                case 0x01CB:
                                    // Line Style : lineWidth
                                    //@link : http://msdn.microsoft.com/en-us/library/dd926964(v=office.12).aspx
                                    $arrShpPrimaryOpt['lineWidth'] = \PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']);
                                    break;
                                case 0x01D6:
                                    // Line Style : lineJoinStyle
                                    //@link : http://msdn.microsoft.com/en-us/library/dd909643(v=office.12).aspx
                                    break;
                                case 0x01D7:
                                    // Line Style : lineEndCapStyle
                                    //@link : http://msdn.microsoft.com/en-us/library/dd925071(v=office.12).aspx
                                    break;
                                case 0x01FF:
                                    // Line Style : Line Style Boolean Properties
                                    //@link : http://msdn.microsoft.com/en-us/library/dd951605(v=office.12).aspx
                                    break;
                                case 0x0201:
                                    // Shadow Style : shadowColor
                                    //@link : http://msdn.microsoft.com/en-us/library/dd923454(v=office.12).aspx
                                    break;
                                case 0x0204:
                                    // Shadow Style : shadowOpacity
                                    //@link : http://msdn.microsoft.com/en-us/library/dd920720(v=office.12).aspx
                                    break;
                                case 0x0205:
                                    // Shadow Style : shadowOffsetX
                                    //@link : http://msdn.microsoft.com/en-us/library/dd945280(v=office.12).aspx
                                    $arrShpPrimaryOpt['shadowOffsetX'] = \PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']);
                                    break;
                                case 0x0206:
                                    // Shadow Style : shadowOffsetY
                                    //@link : http://msdn.microsoft.com/en-us/library/dd907855(v=office.12).aspx
                                    $arrShpPrimaryOpt['shadowOffsetY'] = \PhpOffice\PhpPowerpoint\Shared\Drawing::emuToPixels($opt['op']);
                                    break;
                                case 0x023F:
                                    // Shadow Style : Shadow Style Boolean Properties
                                    //@link : http://msdn.microsoft.com/en-us/library/dd947887(v=office.12).aspx
                                    break;
                                case 0x0304:
                                    // Shape : bWMode
                                    //@link : http://msdn.microsoft.com/en-us/library/dd947659(v=office.12).aspx
                                    break;
                                case 0x033F:
                                    // Shape Boolean Properties
                                    //@link : http://msdn.microsoft.com/en-us/library/dd951345(v=office.12).aspx
                                    break;
                                default:
                                    throw new \Exception('Feature not implemented (l.'.__LINE__.' : '.dechex($opt['opid'].')'));
                                    break;
                            }
                        }
                        $pos += $shapePrimaryOptions['recLen'];
                        $rHeader['recLen'] -= $shapePrimaryOptions['recLen'];
                    }
                    
                    $shpSecondaryOptions1 = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    $bShpSecondaryOpt1 = false;
                    if ($shpSecondaryOptions1['recVer'] == 0x3 && $shpSecondaryOptions1['recType'] == 0xF121) {
                        $pos += 8;
                        $rHeader['recLen'] -= 8;
                        $bShpSecondaryOpt1 = true;
                        // echo '@$rhShapeSecondaryOptions1'.EOL;
                    }
                    
                    $shpTertiaryOptions1 = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    $bShpTertiaryOptions1 = false;
                    if ($shpTertiaryOptions1['recVer'] == 0x3 && $shpTertiaryOptions1['recType'] == 0xF122) {
                        $pos += 8;
                        $rHeader['recLen'] -= 8;
                        $bShpTertiaryOptions1 = true;
                        // echo '@$rhShapeTertiaryOptions1'.EOL;
                    }
                    
                    $rhChildAnchor = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    if ($rhChildAnchor['recVer'] == 0x0 && $rhChildAnchor['recInstance'] == 0x000 && $rhChildAnchor['recType'] == 0xF00F && $rhChildAnchor['recLen'] == 0x00000010) {
                        $pos += 8;
                        $rHeader['recLen'] -= 8;
                        // echo '@$rhChildAnchor'.EOL;
                    }
                    
                    $rhClientAnchor = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    $pos += 8;
                    $rHeader['recLen'] -= 8;
                    //@link : http://msdn.microsoft.com/en-us/library/dd922797(v=office.12).aspx
                    if ($rhClientAnchor['recVer'] == 0x0 && $rhClientAnchor['recInstance'] == 0x000 && $rhClientAnchor['recType'] == 0xF010 && ($rhClientAnchor['recLen'] == 0x00000008 || $rhClientAnchor['recLen'] == 0x00000010)) {
                        // echo '$rhClientAnchor'.EOL;
                        switch ($rhClientAnchor['recLen']) {
                            case 0x00000008:
                                // echo '$rhClientAnchor:0x000000008'.EOL;
                                $arrClientAnchor['top'] = (int) self::getInt2d($this->streamPowerpointDocument, $pos) / 6;
                                $arrClientAnchor['left'] = (int) self::getInt2d($this->streamPowerpointDocument, $pos + 2) / 6;
                                $arrClientAnchor['width'] = ((int) self::getInt2d($this->streamPowerpointDocument, $pos + 4) / 6) - $arrClientAnchor['left'];
                                $arrClientAnchor['height'] = ((int) self::getInt2d($this->streamPowerpointDocument, $pos + 6) / 6) - $arrClientAnchor['left'];
                                // print_r($arrClientAnchor);
                                $pos += 8;
                                break;
                            case 0x00000010:
                                // echo '@$rhClientAnchor:0x00000010'.EOL;
                                break;
                        }
                    } else {
                        $pos -= 8;
                        $rHeader['recLen'] += 8;
                    }
                    
                    //@link : http://msdn.microsoft.com/en-us/library/dd950927(v=office.12).aspx
                    $clientData = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    $pos += 8;
                    $rHeader['recLen'] -= 8;
                    if ($clientData['recVer'] == 0xF && $clientData['recInstance'] == 0x000 && $clientData['recType'] == 0xF011) {
                        // echo '@$clientData'.EOL;
                    } else {
                        $pos -= 8;
                        $rHeader['recLen'] += 8;
                    }
                    
                    //@link : http://msdn.microsoft.com/en-us/library/dd910958(v=office.12).aspx
                    $clientTextbox = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    $pos += 8;
                    $rHeader['recLen'] -= 8;
                    if ($clientTextbox['recVer'] == 0xF && $clientTextbox['recInstance'] == 0x000 && $clientTextbox['recType'] == 0xF00D) {
                        // echo '@$clientTextbox'.EOL;
                        $strLen = 0;
                        do {
                            $rhRgChildRec = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                            $pos += 8;
                            $rHeader['recLen'] -= 8;
                            $clientTextbox['recLen'] -= 8;
                            switch ($rhRgChildRec['recType']) {
                                case self::RT_INTERACTIVEINFO:
                                    // echo '$clientTextbox:RT_InteractiveInfo'.EOL;
                                    //@link : http://msdn.microsoft.com/en-us/library/dd948623(v=office.12).aspx
                                    if ($rhRgChildRec['recInstance'] == 0x0000) {
                                        //@link : http://msdn.microsoft.com/en-us/library/dd952348(v=office.12).aspx
                                        $rhInteractiveAtom = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                                        $pos += 8;
                                        $rHeader['recLen'] -= 8;
                                        $clientTextbox['recLen'] -= 8;
                                        if ($rhInteractiveAtom['recVer'] != 0x0 || $rhInteractiveAtom['recInstance'] != 0x000 || $rhInteractiveAtom['recType'] != self::RT_INTERACTIVEINFOATOM || $rhInteractiveAtom['recLen'] != 0x00000010) {
                                            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > drawing > OfficeArtDg > groupShape > rgfb > shapePrimaryOptions > clientTextbox).');
                                        }
                                        // soundIdRef
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        // exHyperlinkIdRef
                                        $exHyperlinkIdRef = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        // action
                                        $pos += 1;
                                        $rHeader['recLen'] -= 1;
                                        $clientTextbox['recLen'] -= 1;
                                        // oleVerb
                                        $pos += 1;
                                        $rHeader['recLen'] -= 1;
                                        $clientTextbox['recLen'] -= 1;
                                        // jump
                                        $pos += 1;
                                        $rHeader['recLen'] -= 1;
                                        $clientTextbox['recLen'] -= 1;
                                        // fAnimated (1 bit)
                                        // fStopSound (1 bit)
                                        // fCustomShowReturn (1 bit)
                                        // fVisited (1 bit)
                                        // reserved (4 bits)
                                        $pos += 1;
                                        $rHeader['recLen'] -= 1;
                                        $clientTextbox['recLen'] -= 1;
                                        // hyperlinkType
                                        $pos += 1;
                                        $rHeader['recLen'] -= 1;
                                        $clientTextbox['recLen'] -= 1;
                                        // unused
                                        $pos += 3;
                                        $rHeader['recLen'] -= 3;
                                        $clientTextbox['recLen'] -= 3;
                                        
                                        // Shape
                                        $arrClientTextBox['hyperlink'][]['id'] = $exHyperlinkIdRef;
                                    }
                                    if ($rhRgChildRec['recInstance'] == 0x0001) {
                                        // echo '@todo l.'.__LINE__;
                                    }
                                    break;
                                case self::RT_STYLETEXTPROPATOM:
                                    // echo '$clientTextbox:RT_StyleTextPropAtom'.EOL;
                                    // @link : http://msdn.microsoft.com/en-us/library/dd950647(v=office.12).aspx
                                    $strLenRT = $strLen + 1;
                                    do {
                                        // rgTextPFRun
                                        $countRgTextPFRun = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $strLenRT -= $countRgTextPFRun;
                                        $arrClientTextBox['numTexts']++;
                                        $arrClientTextBox['text'.$arrClientTextBox['numTexts']] = array();
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        
                                        // indent
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                        
                                        $masks = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        
                                        $masksData = array();
                                        $masksData['hasBullet'] = ($masks >> 0) & bindec('1');
                                        $masksData['bulletHasFont'] = ($masks >> 1) & bindec('1');
                                        $masksData['bulletHasColor'] = ($masks >> 2) & bindec('1');
                                        $masksData['bulletHasSize'] = ($masks >> 3) & bindec('1');
                                        $masksData['bulletFont'] = ($masks >> 4) & bindec('1');
                                        $masksData['bulletColor'] = ($masks >> 5) & bindec('1');
                                        $masksData['bulletSize'] = ($masks >> 6) & bindec('1');
                                        $masksData['bulletChar'] = ($masks >> 7) & bindec('1');
                                        $masksData['leftMargin'] = ($masks >> 8) & bindec('1');
                                        $masksData['unused'] = ($masks >> 9) & bindec('1');
                                        $masksData['indent'] = ($masks >> 10) & bindec('1');
                                        $masksData['align'] = ($masks >> 11) & bindec('1');
                                        $masksData['lineSpacing'] = ($masks >> 12) & bindec('1');
                                        $masksData['spaceBefore'] = ($masks >> 13) & bindec('1');
                                        $masksData['spaceAfter'] = ($masks >> 14) & bindec('1');
                                        $masksData['defaultTabSize'] = ($masks >> 15) & bindec('1');
                                        $masksData['fontAlign'] = ($masks >> 16) & bindec('1');
                                        $masksData['charWrap'] = ($masks >> 17) & bindec('1');
                                        $masksData['wordWrap'] = ($masks >> 18) & bindec('1');
                                        $masksData['overflow'] = ($masks >> 19) & bindec('1');
                                        $masksData['tabStops'] = ($masks >> 20) & bindec('1');
                                        $masksData['textDirection'] = ($masks >> 21) & bindec('1');
                                        $masksData['reserved1'] = ($masks >> 22) & bindec('1');
                                        $masksData['bulletBlip'] = ($masks >> 23) & bindec('1');
                                        $masksData['bulletScheme'] = ($masks >> 24) & bindec('1');
                                        $masksData['bulletHasScheme'] = ($masks >> 25) & bindec('1');
    
                                        $bulletFlags = array();
                                        if ($masksData['hasBullet'] == 1 || $masksData['bulletHasFont'] == 1  || $masksData['bulletHasColor'] == 1  || $masksData['bulletHasSize '] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            
                                            $bulletFlags['fHasBullet'] = ($data >> 0) & bindec('1');
                                            $bulletFlags['fBulletHasFont'] = ($data >> 1) & bindec('1');
                                            $bulletFlags['fBulletHasColor'] = ($data >> 2) & bindec('1');
                                            $bulletFlags['fBulletHasSize'] = ($data >> 3) & bindec('1');
                                        }
                                        if ($masksData['bulletChar'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            $arrClientTextBox['text'.$arrClientTextBox['numTexts']]['bulletChar'] = chr($data);
                                        }
                                        if ($masksData['bulletFont'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['bulletSize'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['bulletColor'] == 1) {
                                            $red = self::getInt1d($this->streamPowerpointDocument, $pos);
                                            $pos += 1;
                                            $rHeader['recLen'] -= 1;
                                            $clientTextbox['recLen'] -= 1;
                                            $green = self::getInt1d($this->streamPowerpointDocument, $pos);
                                            $pos += 1;
                                            $rHeader['recLen'] -= 1;
                                            $clientTextbox['recLen'] -= 1;
                                            $blue = self::getInt1d($this->streamPowerpointDocument, $pos);
                                            $pos += 1;
                                            $rHeader['recLen'] -= 1;
                                            $clientTextbox['recLen'] -= 1;
                                            $index = self::getInt1d($this->streamPowerpointDocument, $pos);
                                            $pos += 1;
                                            $rHeader['recLen'] -= 1;
                                            $clientTextbox['recLen'] -= 1;
                                            
                                            if ($index == 0xFE) {
                                                $strColor = str_pad(dechex($red), 2, STR_PAD_LEFT, '0');
                                                $strColor .= str_pad(dechex($green), 2, STR_PAD_LEFT, '0');
                                                $strColor .= str_pad(dechex($blue), 2, STR_PAD_LEFT, '0');
                                            }
                                        }
                                        if ($masksData['align'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            switch ($data) {
                                                case 0x0000:
                                                    $arrClientTextBox['alignH'] = Alignment::HORIZONTAL_LEFT;
                                                    break;
                                                case 0x0001:
                                                    $arrClientTextBox['alignH'] = Alignment::HORIZONTAL_CENTER;
                                                    break;
                                                case 0x0002:
                                                    $arrClientTextBox['alignH'] = Alignment::HORIZONTAL_RIGHT;
                                                    break;
                                                case 0x0003:
                                                    $arrClientTextBox['alignH'] = Alignment::HORIZONTAL_JUSTIFY;
                                                    break;
                                                case 0x0004:
                                                    $arrClientTextBox['alignH'] = Alignment::HORIZONTAL_DISTRIBUTED;
                                                    break;
                                                case 0x0005:
                                                    $arrClientTextBox['alignH'] = Alignment::HORIZONTAL_DISTRIBUTED;
                                                    break;
                                                case 0x0006:
                                                    $arrClientTextBox['alignH'] = Alignment::HORIZONTAL_JUSTIFY;
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                        if ($masksData['lineSpacing'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['spaceBefore'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['spaceAfter'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['leftMargin'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            $arrClientTextBox['text'.$arrClientTextBox['numTexts']]['leftMargin'] = (int)round($data/6);
                                        }
                                        if ($masksData['indent'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            $arrClientTextBox['text'.$arrClientTextBox['numTexts']]['indent'] = (int)round($data/6);
                                        }
                                        if ($masksData['defaultTabSize'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['tabStops'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                        if ($masksData['fontAlign'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['charWrap'] == 1 || $masksData['wordWrap'] == 1 || $masksData['overflow'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['textDirection'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                    } while ($strLenRT > 0);
                                    
                                    $strLenRT = $strLen + 1;
                                    do {
                                        // rgTextCFRun
                                        $countRgTextCFRun = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $strLenRT -= $countRgTextCFRun;
                                        $arrClientTextBox['numParts']++;
                                        $arrClientTextBox['part'.$arrClientTextBox['numParts']] = array();
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        $arrClientTextBox['part'.$arrClientTextBox['numParts']]['length'] = $countRgTextCFRun;
                                        
                                        $masks = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        
                                        $masksData = array();
                                        $masksData['bold'] = ($masks >> 0) & bindec('1');
                                        $masksData['italic'] = ($masks >> 1) & bindec('1');
                                        $masksData['underline'] = ($masks >> 2) & bindec('1');
                                        $masksData['unused1'] = ($masks >> 3) & bindec('1');
                                        $masksData['shadow'] = ($masks >> 4) & bindec('1');
                                        $masksData['fehint'] = ($masks >> 5) & bindec('1');
                                        $masksData['unused2'] = ($masks >> 6) & bindec('1');
                                        $masksData['kumi'] = ($masks >> 7) & bindec('1');
                                        $masksData['unused3'] = ($masks >> 8) & bindec('1');
                                        $masksData['emboss'] = ($masks >> 9) & bindec('1');
                                        $masksData['fHasStyle'] = ($masks >> 10) & bindec('1111');
                                        $masksData['unused4'] = ($masks >> 14) & bindec('11');
                                        $masksData['typeface'] = ($masks >> 16) & bindec('1');
                                        $masksData['size'] = ($masks >> 17) & bindec('1');
                                        $masksData['color'] = ($masks >> 18) & bindec('1');
                                        $masksData['position'] = ($masks >> 19) & bindec('1');
                                        $masksData['pp10ext'] = ($masks >> 20) & bindec('1');
                                        $masksData['oldEATypeface'] = ($masks >> 21) & bindec('1');
                                        $masksData['ansiTypeface'] = ($masks >> 22) & bindec('1');
                                        $masksData['symbolTypeface'] = ($masks >> 23) & bindec('1');
                                        $masksData['newEATypeface'] = ($masks >> 24) & bindec('1');
                                        $masksData['csTypeface'] = ($masks >> 25) & bindec('1');
                                        $masksData['pp11ext'] = ($masks >> 26) & bindec('1');
                                        if ($masksData['bold'] == 1 || $masksData['italic'] == 1 || $masksData['underline'] == 1 || $masksData['shadow'] == 1 || $masksData['fehint'] == 1 ||  $masksData['kumi'] == 1 ||  $masksData['emboss'] == 1 ||  $masksData['fHasStyle'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                             
                                            $fontStyleFlags['bold'] = ($data >> 0) & bindec('1');
                                            $fontStyleFlags['italic'] = ($data >> 1) & bindec('1');
                                            $fontStyleFlags['underline'] = ($data >> 2) & bindec('1');
                                            $fontStyleFlags['unused1'] = ($data >> 3) & bindec('1');
                                            $fontStyleFlags['shadow'] = ($data >> 4) & bindec('1');
                                            $fontStyleFlags['fehint'] = ($data >> 5) & bindec('1');
                                            $fontStyleFlags['unused2'] = ($data >> 6) & bindec('1');
                                            $fontStyleFlags['kumi'] = ($data >> 7) & bindec('1');
                                            $fontStyleFlags['unused3'] = ($data >> 8) & bindec('1');
                                            $fontStyleFlags['emboss'] = ($data >> 9) & bindec('1');
                                            $fontStyleFlags['pp9rt'] = ($data >> 10) & bindec('1111');
                                            $fontStyleFlags['unused4'] = ($data >> 14) & bindec('11');
                                            
                                            $arrClientTextBox['part'.$arrClientTextBox['numParts']]['bold'] = ($fontStyleFlags['bold'] == 1) ? true : false;
                                            $arrClientTextBox['part'.$arrClientTextBox['numParts']]['italic'] = ($fontStyleFlags['italic'] == 1) ? true : false;
                                            $arrClientTextBox['part'.$arrClientTextBox['numParts']]['underline'] = ($fontStyleFlags['underline'] == 1) ? true : false;
                                        }
                                        if ($masksData['typeface'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            $arrClientTextBox['part'.$arrClientTextBox['numParts']]['fontName'] = isset($this->arrayFonts[$data]) ? $this->arrayFonts[$data] : '';
                                        }
                                        if ($masksData['oldEATypeface'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                        if ($masksData['ansiTypeface'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                        if ($masksData['symbolTypeface'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                        if ($masksData['size'] == 1) {
                                            $arrClientTextBox['part'.$arrClientTextBox['numParts']]['fontSize'] = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['color'] == 1) {
                                            $red = self::getInt1d($this->streamPowerpointDocument, $pos);
                                            $pos += 1;
                                            $rHeader['recLen'] -= 1;
                                            $clientTextbox['recLen'] -= 1;
                                            $green = self::getInt1d($this->streamPowerpointDocument, $pos);
                                            $pos += 1;
                                            $rHeader['recLen'] -= 1;
                                            $clientTextbox['recLen'] -= 1;
                                            $blue = self::getInt1d($this->streamPowerpointDocument, $pos);
                                            $pos += 1;
                                            $rHeader['recLen'] -= 1;
                                            $clientTextbox['recLen'] -= 1;
                                            $index = self::getInt1d($this->streamPowerpointDocument, $pos);
                                            $pos += 1;
                                            $rHeader['recLen'] -= 1;
                                            $clientTextbox['recLen'] -= 1;
                                            
                                            if ($index == 0xFE) {
                                                $strColor = str_pad(dechex($red), 2, STR_PAD_LEFT, '0');
                                                $strColor .= str_pad(dechex($green), 2, STR_PAD_LEFT, '0');
                                                $strColor .= str_pad(dechex($blue), 2, STR_PAD_LEFT, '0');
                                                
                                                $arrClientTextBox['part'.$arrClientTextBox['numParts']]['color'] = new Color('FF'.$strColor);
                                            }
                                        }
                                        if ($masksData['position'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                    } while ($strLenRT > 0);
                                    break;
                                case self::RT_TEXTCHARSATOM:
                                    // echo '$clientTextbox:RT_TextCharsAtom'.EOL;
                                    // @link : http://msdn.microsoft.com/en-us/library/dd772921(v=office.12).aspx
                                    $strLen = (int)($rhRgChildRec['recLen']/2);
                                    for ($inc = 0; $inc < $rhRgChildRec['recLen']/2; $inc++) {
                                        $char = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        if ($char == 0x0B) {
                                            $char = 0x20;
                                        }
                                        $arrClientTextBox['text'] .= chr($char);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    // echo $arrClientTextBox['text'].EOL;
                                    break;
                                case self::RT_TEXTHEADERATOM:
                                    // echo '$clientTextbox:RT_TextHeaderAtom'.EOL;
                                    // @link : http://msdn.microsoft.com/en-us/library/dd905272(v=office.12).aspx
                                    // textType
                                    $pos += 4;
                                    $rHeader['recLen'] -= 4;
                                    $clientTextbox['recLen'] -= 4;
                                    break;
                                case self::RT_TEXTINTERACTIVEINFOATOM:
                                    // echo '$clientTextbox:RT_TextInteractiveInfoAtom'.EOL;
                                    //@link : http://msdn.microsoft.com/en-us/library/dd947973(v=office.12).aspx
                                    if ($rhRgChildRec['recInstance'] == 0x0000) {
                                        //@link : http://msdn.microsoft.com/en-us/library/dd944072(v=office.12).aspx
                                        $arrClientTextBox['hyperlink'][count($arrClientTextBox['hyperlink']) - 1]['start'] = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        
                                        $arrClientTextBox['hyperlink'][count($arrClientTextBox['hyperlink']) - 1]['end'] = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                    }
                                    if ($rhRgChildRec['recInstance'] == 0x0001) {
                                        // echo '@todo l.'.__LINE__;
                                    }
                                    break;
                                case self::RT_TEXTSPECIALINFOATOM:
                                    // echo '$clientTextbox:RT_TextSpecialInfoAtom'.EOL;
                                    // @link : http://msdn.microsoft.com/en-us/library/dd945296(v=office.12).aspx
                                    $strLenRT = $strLen + 1;
                                    do {
                                        $count = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        $strLenRT -= $count;
                                        $data = self::getInt4d($this->streamPowerpointDocument, $pos);
                                        $pos += 4;
                                        $rHeader['recLen'] -= 4;
                                        $clientTextbox['recLen'] -= 4;
                                        $masksData = array();
                                        $masksData['spell'] = ($data >> 0) & bindec('1');
                                        $masksData['lang'] = ($data >> 1) & bindec('1');
                                        $masksData['altLang'] = ($data >> 2) & bindec('1');
                                        $masksData['unused1'] = ($data >> 3) & bindec('1');
                                        $masksData['unused2'] = ($data >> 4) & bindec('1');
                                        $masksData['fPp10ext'] = ($data >> 5) & bindec('1');
                                        $masksData['fBidi'] = ($data >> 6) & bindec('1');
                                        $masksData['unused3'] = ($data >> 7) & bindec('1');
                                        $masksData['reserved1'] = ($data >> 8) & bindec('1');
                                        $masksData['smartTag'] = ($data >> 9) & bindec('1');
                                        
                                        if ($masksData['spell'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            $masksSpell = array();
                                            $masksSpell['error'] = ($data >> 0) & bindec('1');
                                            $masksSpell['clean'] = ($data >> 1) & bindec('1');
                                            $masksSpell['grammar'] = ($data >> 2) & bindec('1');
                                        }
                                        if ($masksData['lang'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['altLang'] == 1) {
                                            $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                        }
                                        if ($masksData['fBidi'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                        if ($masksData['fPp10ext'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                        if ($masksData['smartTag'] == 1) {
                                            // echo '@todo l.'.__LINE__;
                                        }
                                    } while ($strLenRT > 0);
                                    break;
                                case self::RT_TEXTRULERATOM:
                                    // echo '$clientTextbox:RT_TextRulerAtom'.EOL;
                                    // @link : http://msdn.microsoft.com/en-us/library/dd953212(v=office.12).aspx
                                    $data = self::getInt4d($this->streamPowerpointDocument, $pos);
                                    $pos += 4;
                                    $rHeader['recLen'] -= 4;
                                    $clientTextbox['recLen'] -= 4;
                                    $masksData = array();
                                    $masksData['fDefaultTabSize'] = ($data >> 0) & bindec('1');
                                    $masksData['fCLevels'] = ($data >> 1) & bindec('1');
                                    $masksData['fTabStops'] = ($data >> 2) & bindec('1');
                                    $masksData['fLeftMargin1'] = ($data >> 3) & bindec('1');
                                    $masksData['fLeftMargin2'] = ($data >> 4) & bindec('1');
                                    $masksData['fLeftMargin3'] = ($data >> 5) & bindec('1');
                                    $masksData['fLeftMargin4'] = ($data >> 6) & bindec('1');
                                    $masksData['fLeftMargin5'] = ($data >> 7) & bindec('1');
                                    $masksData['fIndent1'] = ($data >> 8) & bindec('1');
                                    $masksData['fIndent2'] = ($data >> 9) & bindec('1');
                                    $masksData['fIndent3'] = ($data >> 10) & bindec('1');
                                    $masksData['fIndent4'] = ($data >> 11) & bindec('1');
                                    $masksData['fIndent5'] = ($data >> 12) & bindec('1');
                                    
                                    if ($masksData['fCLevels'] == 1) {
                                        // echo '@todo l.'.__LINE__;
                                    }
                                    if ($masksData['fDefaultTabSize'] == 1) {
                                        // echo '@todo l.'.__LINE__;
                                    }
                                    if ($masksData['fTabStops'] == 1) {
                                        $count = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                        $arrayTabStops = array();
                                        for ($inc = 0; $inc < $count; $inc++) {
                                            $position = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            $type = self::getInt2d($this->streamPowerpointDocument, $pos);
                                            $pos += 2;
                                            $rHeader['recLen'] -= 2;
                                            $clientTextbox['recLen'] -= 2;
                                            $arrayTabStops[] = array(
                                                'position' => $position,
                                                'type' => $type,
                                            );
                                        }
                                    }
                                    if ($masksData['fLeftMargin1'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fIndent1'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fLeftMargin2'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fIndent2'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fLeftMargin3'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fIndent3'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fLeftMargin4'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fIndent4'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fLeftMargin5'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    if ($masksData['fIndent5'] == 1) {
                                        $data = self::getInt2d($this->streamPowerpointDocument, $pos);
                                        $pos += 2;
                                        $rHeader['recLen'] -= 2;
                                        $clientTextbox['recLen'] -= 2;
                                    }
                                    break;
                                default:
                                    // echo EOL;
                                    // print_r($rhRgChildRec);
                                    // echo EOL;
                                    // echo '@$type : '.dechex($rhRgChildRec['recType']).EOL;
                            }
                        } while ($clientTextbox['recLen'] > 0);
                    } else {
                        $pos -= 8;
                        $rHeader['recLen'] += 8;
                    }
                    $shpSecondaryOptions2 = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    $pos += 8;
                    $rHeader['recLen'] -= 8;
                    if ($bShpSecondaryOpt1 == true && $shpSecondaryOptions2['recVer'] == 0x3 && $shpSecondaryOptions2['recType'] == 0xF121) {
                    } else {
                        $pos -= 8;
                        $rHeader['recLen'] += 8;
                    }
                    
                    $shpTertiaryOptions2 = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    $pos += 8;
                    $rHeader['recLen'] -= 8;
                    if ($bShpTertiaryOptions1 == true && $shpTertiaryOptions2['recVer'] == 0x3 && $shpTertiaryOptions2['recType'] == 0xF122) {
                    } else {
                        $pos -= 8;
                        $rHeader['recLen'] += 8;
                    }
                    break;
            }
            
            // Cleaning variables
            if (empty($arrClientTextBox['hyperlink'])) {
                unset($arrClientTextBox['hyperlink']);
            }
            if (empty($arrClientTextBox['text'])) {
                unset($arrClientTextBox['text']);
            }
            if (empty($arrClientTextBox['numParts'])) {
                unset($arrClientTextBox['numParts']);
            }
            if (empty($arrClientTextBox['numTexts'])) {
                unset($arrClientTextBox['numTexts']);
            }
            // echo EOL;
            if (isset($arrShpPrimaryOpt['pib']) && isset($this->arrayPictures[$arrShpPrimaryOpt['pib'] - 1])) {
                // echo '//IMAGE'.EOL;
                $gdImage = imagecreatefromstring($this->arrayPictures[$arrShpPrimaryOpt['pib'] - 1]);
                
                $shape = new MemoryDrawing();
                $shape->setImageResource($gdImage);
                
                if (isset($arrShpPrimaryOpt['shadowOffsetX']) && $arrShpPrimaryOpt['shadowOffsetX'] != 0 && isset($arrShpPrimaryOpt['shadowOffsetY']) && $arrShpPrimaryOpt['shadowOffsetY'] != 0) {
                    $shape->getShadow()->setVisible(true);
                    if ($arrShpPrimaryOpt['shadowOffsetX'] > 0 && $arrShpPrimaryOpt['shadowOffsetX'] == $arrShpPrimaryOpt['shadowOffsetY']) {
                        $shape->getShadow()->setDistance($arrShpPrimaryOpt['shadowOffsetX'])->setDirection(45);
                    }
                }
                if (!is_null($shape) && !empty($arrClientAnchor)) {
                    $shape->setOffsetX($arrClientAnchor['left']);
                    $shape->setOffsetY($arrClientAnchor['top']);
                    $shape->setWidth($arrClientAnchor['width']);
                    $shape->setHeight($arrClientAnchor['height']);
                }
            }
            
            if (!empty($arrClientTextBox) && isset($arrClientTextBox)) {
                // echo '//TEXT'.EOL;
                // echo '<pre>';
                // print_r($arrClientTextBox);
                // echo '</pre>';
                $shape = new RichText();
                if (isset($arrClientTextBox['alignH'])) {
                    $shape->getActiveParagraph()->getAlignment()->setHorizontal($arrClientTextBox['alignH']);
                }
                
                $start = 0;
                // echo $arrClientTextBox['text'].EOL;
                
                $lastLevel = -1;
                $lastMarginLeft = 0;
                for ($inc = 1; $inc <= $arrClientTextBox['numParts']; $inc++) {
                    if ($arrClientTextBox['numParts'] == $arrClientTextBox['numTexts'] && isset($arrClientTextBox['text'.$inc])) {
                        if (isset($arrClientTextBox['text'.$inc]['bulletChar'])) {
                            $shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
                            $shape->getActiveParagraph()->getBulletStyle()->setBulletChar($arrClientTextBox['text'.$inc]['bulletChar']);
                        }
                        // Indent
                        $indent = 0;
                        if (isset($arrClientTextBox['text'.$inc]['indent'])) {
                            $indent = $arrClientTextBox['text'.$inc]['indent'];
                        }
                        if (isset($arrClientTextBox['text'.$inc]['leftMargin'])) {
                            if ($lastMarginLeft > $arrClientTextBox['text'.$inc]['leftMargin']) {
                                $lastLevel--;
                            }
                            if ($lastMarginLeft < $arrClientTextBox['text'.$inc]['leftMargin']) {
                                $lastLevel++;
                            }
                            $shape->getActiveParagraph()->getAlignment()->setLevel($lastLevel);
                            $lastMarginLeft = $arrClientTextBox['text'.$inc]['leftMargin'];
                            
                            $shape->getActiveParagraph()->getAlignment()->setMarginLeft($arrClientTextBox['text'.$inc]['leftMargin']);
                            $shape->getActiveParagraph()->getAlignment()->setIndent($indent - $arrClientTextBox['text'.$inc]['leftMargin']);
                        }
                    
                    }
                    // Texte
                    $sText = substr($arrClientTextBox['text'], $start, $arrClientTextBox['part'.$inc]['length']);
                    $sHyperlinkURL = '';
                    if (empty($sText)) {
                        // Is there a hyperlink ?
                        if (!empty($arrClientTextBox['hyperlink'])) {
                            foreach ($arrClientTextBox['hyperlink'] as $itmHyperlink) {
                                if ($itmHyperlink['start'] == $start && ($itmHyperlink['end'] - $itmHyperlink['start']) == $arrClientTextBox['part'.$inc]['length']) {
                                    $sText = $this->arrayHyperlinks[$itmHyperlink['id']]['text'];
                                    $sHyperlinkURL = $this->arrayHyperlinks[$itmHyperlink['id']]['url'];
                                    break;
                                }
                            }
                        }
                    }
                    // New paragraph
                    $bCreateParagraph = false;
                    if (strpos($sText, "\r") != false) {
                        $bCreateParagraph = true;
                        $sText = str_replace("\r", '', $sText);
                    }
                    // TextRun
                    $txtRun = $shape->createTextRun($sText);
                    if (isset($arrClientTextBox['part'.$inc]['bold'])) {
                        $txtRun->getFont()->setBold($arrClientTextBox['part'.$inc]['bold']);
                    }
                    if (isset($arrClientTextBox['part'.$inc]['italic'])) {
                        $txtRun->getFont()->setItalic($arrClientTextBox['part'.$inc]['italic']);
                    }
                    if (isset($arrClientTextBox['part'.$inc]['underline'])) {
                        $txtRun->getFont()->setUnderline($arrClientTextBox['part'.$inc]['underline']);
                    }
                    if (isset($arrClientTextBox['part'.$inc]['fontName'])) {
                        $txtRun->getFont()->setName($arrClientTextBox['part'.$inc]['fontName']);
                    }
                    if (isset($arrClientTextBox['part'.$inc]['fontSize'])) {
                        $txtRun->getFont()->setSize($arrClientTextBox['part'.$inc]['fontSize']);
                    }
                    if (isset($arrClientTextBox['part'.$inc]['color'])) {
                        $txtRun->getFont()->setColor($arrClientTextBox['part'.$inc]['color']);
                    }
                    // Hyperlink
                    if (!empty($sHyperlinkURL)) {
                        $txtRun->setHyperlink(new Hyperlink($sHyperlinkURL));
                    }
                    
                    $start += $arrClientTextBox['part'.$inc]['length'];
                    if ($bCreateParagraph) {
                        $shape->createParagraph();
                    }
                }
                if (!is_null($shape) && !empty($arrClientAnchor)) {
                    $shape->setOffsetX($arrClientAnchor['left']);
                    $shape->setOffsetY($arrClientAnchor['top']);
                    $shape->setWidth($arrClientAnchor['width']);
                    $shape->setHeight($arrClientAnchor['height']);
                }
            }
            if (isset($arrShpPrimaryOpt['line']) && $arrShpPrimaryOpt['line'] == true) {
                // echo '//LINE'.EOL;
                $shape = new Line(0, 0, 0, 0);
                if (isset($arrShpPrimaryOpt['lineColor'])) {
                    $shape->getBorder()->getColor()->setARGB('FF'.$arrShpPrimaryOpt['lineColor']);
                }
                if (!empty($arrClientAnchor)) {
                    $shape->setOffsetX($arrClientAnchor['left']);
                    $shape->setOffsetY($arrClientAnchor['top']);
                    $shape->setWidth($arrClientAnchor['width']);
                    $shape->setHeight($arrShpPrimaryOpt['lineWidth']);
                }
            }
            
            if (isset($arrShpPrimaryOpt['insetBottom'])) {
                $shape->setInsetBottom($arrShpPrimaryOpt['insetBottom']);
            }
            if (isset($arrShpPrimaryOpt['insetLeft'])) {
                $shape->setInsetLeft($arrShpPrimaryOpt['insetLeft']);
            }
            if (isset($arrShpPrimaryOpt['insetRight'])) {
                $shape->setInsetRight($arrShpPrimaryOpt['insetRight']);
            }
            if (isset($arrShpPrimaryOpt['insetTop'])) {
                $shape->setInsetTop($arrShpPrimaryOpt['insetTop']);
            }
            
            if (!is_null($shape) && $shape instanceof AbstractShape) {
                // echo '//SHAPE'.EOL;
                $oSlide->addShape($shape);
            }
            // echo '//END.....'.EOL.EOL.EOL;
        } while ($rHeader['recLen'] > 0);
        
        
        // *** slideSchemeColorSchemeAtom (40 bytes)
        // slideSchemeColorSchemeAtom > rh
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0x0 || $rHeader['recInstance'] != 0x001 || $rHeader['recType'] != self::RT_COLORSCHEMEATOM || $rHeader['recLen'] != 0x00000020) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > slideSchemeColorSchemeAtom > rh).');
        }
        // slideSchemeColorSchemeAtom > rgSchemeColor
        $rgSchemeColor = array();
        for ($inc = 0; $inc <= 7; $inc++) {
            $rgSchemeColor[] = array(
                'red' => self::getInt1d($this->streamPowerpointDocument, $pos + $inc * 4),
                'green' => self::getInt1d($this->streamPowerpointDocument, $pos + $inc * 4 + 1),
                'blue' => self::getInt1d($this->streamPowerpointDocument, $pos + $inc * 4 + 2),
            );
        }
        $pos += (8 * 4);
        
        // *** slideNameAtom (variable)
        // *** slideProgTagsContainer (variable).
        // slideProgTagsContainer > rh
        $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0xF || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_PROGTAGS) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTSlide > slideProgTagsContainer > rh).');
        }
    }
    
    /**
     * DocumentContainer
     * @link http://msdn.microsoft.com/en-us/library/dd947357(v=office.12).aspx
     * @param integer $pos
     */
    private function readRTDocument($pos)
    {
        $documentAtom = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        $pos += 8;
        if ($documentAtom['recVer'] != 0x1 || $documentAtom['recInstance'] != 0x000 || $documentAtom['recType'] != self::RT_DOCUMENTATOM) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTDocument > DocumentAtom).');
        }
        $pos += $documentAtom['recLen'];
        
        $exObjList = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        if ($exObjList['recVer'] == 0xF && $exObjList['recInstance'] == 0x000 && $exObjList['recType'] == self::RT_EXTERNALOBJECTLIST) {
            $pos += 8;
            // exObjListAtom > rh
            $exObjListAtom = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
            if ($exObjListAtom['recVer'] != 0x0 || $exObjListAtom['recInstance'] != 0x000 || $exObjListAtom['recType'] != self::RT_EXTERNALOBJECTLISTATOM || $exObjListAtom['recLen'] != 0x00000004) {
                throw new \Exception('File PowerPoint 97 in error (Location : RTDocument > DocumentAtom > exObjList > exObjListAtom).');
            }
            $pos += 8;
            // exObjListAtom > exObjIdSeed
            $pos += 4;
            // rgChildRec
            $exObjList['recLen'] -= 12;
            do {
                $childRec = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                $pos += 8;
                $exObjList['recLen'] -= 8;
                switch ($childRec['recType']) {
                    case self::RT_EXTERNALHYPERLINK:
                        //@link : http://msdn.microsoft.com/en-us/library/dd944995(v=office.12).aspx
                        // exHyperlinkAtom > rh
                        $exHyperlinkAtom = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                        if ($exHyperlinkAtom['recVer'] != 0x0 || $exHyperlinkAtom['recInstance'] != 0x000 || $exHyperlinkAtom['recType'] != self::RT_EXTERNALHYPERLINKATOM || $exObjListAtom['recLen'] != 0x00000004) {
                            throw new \Exception('File PowerPoint 97 in error (Location : RTDocument > DocumentAtom > exObjList > rgChildRec > RT_ExternalHyperlink).');
                        }
                        $pos += 8;
                        $exObjList['recLen'] -= 8;
                        // exHyperlinkAtom > exHyperlinkId
                        $exHyperlinkId = self::getInt4d($this->streamPowerpointDocument, $pos);
                        $pos += 4;
                        $exObjList['recLen'] -= 4;
                        
                        $this->arrayHyperlinks[$exHyperlinkId] = array();
                        // friendlyNameAtom
                        $friendlyNameAtom  = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                        if ($friendlyNameAtom['recVer'] == 0x0 && $friendlyNameAtom['recInstance'] == 0x000 && $friendlyNameAtom['recType'] == self::RT_CSTRING && $friendlyNameAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $this->arrayHyperlinks[$exHyperlinkId]['text'] = '';
                            for ($inc = 0; $inc < ($friendlyNameAtom['recLen'] / 2); $inc++) {
                                $char = self::getInt2d($this->streamPowerpointDocument, $pos);
                                $pos += 2;
                                $exObjList['recLen'] -= 2;
                                $this->arrayHyperlinks[$exHyperlinkId]['text'] .= chr($char);
                            }
                        }
                        // targetAtom
                        $targetAtom  = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                        if ($targetAtom['recVer'] == 0x0 && $targetAtom['recInstance'] == 0x001 && $targetAtom['recType'] == self::RT_CSTRING && $targetAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $this->arrayHyperlinks[$exHyperlinkId]['url'] = '';
                            for ($inc = 0; $inc < ($targetAtom['recLen'] / 2); $inc++) {
                                $char = self::getInt2d($this->streamPowerpointDocument, $pos);
                                $pos += 2;
                                $exObjList['recLen'] -= 2;
                                $this->arrayHyperlinks[$exHyperlinkId]['url'] .= chr($char);
                            }
                        }
                        // locationAtom
                        $locationAtom  = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                        if ($locationAtom['recVer'] == 0x0 && $locationAtom['recInstance'] == 0x003 && $locationAtom['recType'] == self::RT_CSTRING && $locationAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $string = '';
                            for ($inc = 0; $inc < ($locationAtom['recLen'] / 2); $inc++) {
                                $char = self::getInt2d($this->streamPowerpointDocument, $pos);
                                $pos += 2;
                                $exObjList['recLen'] -= 2;
                                $string .= chr($char);
                            }
                        }
                        break;
                    default:
                        throw new \Exception('Feature not implemented (l.'.__LINE__.' : '.dechex($childRec['recType'].')'));
                }
            } while ($exObjList['recLen'] > 0);
        }
        
        //@link : http://msdn.microsoft.com/en-us/library/dd907813(v=office.12).aspx
        $documentTextInfo = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
        if ($documentTextInfo['recVer'] == 0xF && $documentTextInfo['recInstance'] == 0x000 && $documentTextInfo['recType'] == self::RT_ENVIRONMENT) {
            $pos += 8;
            //@link : http://msdn.microsoft.com/en-us/library/dd952717(v=office.12).aspx
            $kinsoku = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
            if ($kinsoku['recVer'] == 0xF && $kinsoku['recInstance'] == 0x002 && $kinsoku['recType'] == self::RT_KINSOKU) {
                $pos += 8;
                $pos += $kinsoku['recLen'];
            }
            
            //@link : http://msdn.microsoft.com/en-us/library/dd948152(v=office.12).aspx
            $fontCollection = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
            if ($fontCollection['recVer'] == 0xF && $fontCollection['recInstance'] == 0x000 && $fontCollection['recType'] == self::RT_FONTCOLLECTION) {
                $pos += 8;
                do {
                    $fontEntityAtom = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    $pos += 8;
                    $fontCollection['recLen'] -= 8;
                    if ($fontEntityAtom['recVer'] != 0x0 || $fontEntityAtom['recInstance'] > 128 || $fontEntityAtom['recType'] != self::RT_FONTENTITYATOM) {
                        throw new \Exception('File PowerPoint 97 in error (Location : RTDocument > RT_Environment > RT_FontCollection > RT_FontEntityAtom).');
                    } else {
                        $string = '';
                        for ($inc = 0; $inc < 32; $inc++) {
                            $char = self::getInt2d($this->streamPowerpointDocument, $pos);
                            $pos += 2;
                            $fontCollection['recLen'] -= 2;
                            $string .= chr($char);
                        }
                        $this->arrayFonts[] = $string;
                        
                        // lfCharSet (1 byte)
                        $pos += 1;
                        $fontCollection['recLen'] -= 1;
                        
                        // fEmbedSubsetted (1 bit)
                        // unused (7 bits)
                        $pos += 1;
                        $fontCollection['recLen'] -= 1;
                        
                        // rasterFontType (1 bit)
                        // deviceFontType (1 bit)
                        // truetypeFontType (1 bit)
                        // fNoFontSubstitution (1 bit)
                        // reserved (4 bits)
                        $pos += 1;
                        $fontCollection['recLen'] -= 1;
                        
                        // lfPitchAndFamily (1 byte)
                        $pos += 1;
                        $fontCollection['recLen'] -= 1;
                    }
                    
                    $fontEmbedData1 = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    if ($fontEmbedData1['recVer'] == 0x0 && $fontEmbedData1['recInstance'] >= 0x000 && $fontEmbedData1['recInstance'] <= 0x003 && $fontEmbedData1['recType'] == self::RT_FONTEMBEDDATABLOB) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData1['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData1['recLen'];
                    }
                    
                    $fontEmbedData2 = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    if ($fontEmbedData2['recVer'] == 0x0 && $fontEmbedData2['recInstance'] >= 0x000 && $fontEmbedData2['recInstance'] <= 0x003 && $fontEmbedData2['recType'] == self::RT_FONTEMBEDDATABLOB) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData2['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData2['recLen'];
                    }
                    
                    $fontEmbedData3 = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    if ($fontEmbedData3['recVer'] == 0x0 && $fontEmbedData3['recInstance'] >= 0x000 && $fontEmbedData3['recInstance'] <= 0x003 && $fontEmbedData3['recType'] == self::RT_FONTEMBEDDATABLOB) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData3['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData3['recLen'];
                    }
                    
                    $fontEmbedData4 = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
                    if ($fontEmbedData4['recVer'] == 0x0 && $fontEmbedData4['recInstance'] >= 0x000 && $fontEmbedData4['recInstance'] <= 0x003 && $fontEmbedData4['recType'] == self::RT_FONTEMBEDDATABLOB) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData4['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData4['recLen'];
                    }
                    
                } while ($fontCollection['recLen'] > 0);
            }
        }
    }
    
    /**
     * Read a record header
     * @param string $stream
     * @param intger $pos
     * @return multitype:boolean Ambigous <number, boolean>
     */
    private function loadRecordHeader($stream, $pos)
    {
        $rec = self::getInt2d($stream, $pos);
        $recType = self::getInt2d($stream, $pos + 2);
        $recLen = self::getInt4d($stream, $pos + 4);
        return array(
            'recVer' => ($rec >> 0) & bindec('1111'),
            'recInstance' => ($rec >> 4) & bindec('111111111111'),
            'recType' => $recType,
            'recLen' => $recLen,
        );
    }
    
    /**
     * Read 8-bit unsigned integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt1d($data, $pos)
    {
        return ord($data[$pos]);
    }
    
    /**
     * Read 16-bit unsigned integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt2d($data, $pos)
    {
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8);
    }
    
    /**
     * Read 32-bit signed integer
     *
     * @param string $data
     * @param int $pos
     * @return int
     */
    public static function getInt4d($data, $pos)
    {
        // FIX: represent numbers correctly on 64-bit system
        // http://sourceforge.net/tracker/index.php?func=detail&aid=1487372&group_id=99160&atid=623334
        // Hacked by Andreas Rehm 2006 to ensure correct result of the <<24 block on 32 and 64bit systems
        $or24 = ord($data[$pos + 3]);
        if ($or24 >= 128) {
            // negative number
            $ord24 = -abs((256 - $or24) << 24);
        } else {
            $ord24 = ($or24 & 127) << 24;
        }
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | $ord24;
    }
}
