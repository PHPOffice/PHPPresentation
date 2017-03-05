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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Reader;

use PhpOffice\Common\Microsoft\OLERead;
use PhpOffice\Common\Text;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\Shape;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Bullet;

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
     * Array with Notes
     */
    private $arrayNotes = array();
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
     * @var int
     */
    private $offsetPersistDirectory;
    /**
     * Output Object
     * @var PhpPresentation
     */
    private $oPhpPresentation;
    /**
     * Group Object
     * @var Group
     */
    private $oCurrentGroup;
    /**
     * @var boolean
     */
    private $bFirstShapeGroup = false;
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
     * @var integer
     */
    private $inMainType;
    /**
     * @var integer
     */
    private $currentNote;

    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     *
     * @param  string $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function canRead($pFilename)
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     *
     * @param  string    $pFilename
     * @throws \Exception
     * @return boolean
     */
    public function fileSupportsUnserializePhpPresentation($pFilename = '')
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
     * Loads PhpPresentation Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     * @throws \Exception
     */
    public function load($pFilename)
    {
        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new \Exception("Invalid file format for PhpOffice\PhpPresentation\Reader\PowerPoint97: " . $pFilename . ".");
        }

        return $this->loadFile($pFilename);
    }

    /**
     * Load PhpPresentation Serialized file
     *
     * @param  string        $pFilename
     * @return \PhpOffice\PhpPresentation\PhpPresentation
     */
    private function loadFile($pFilename)
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();

        // Read OLE Blocks
        $this->loadOLE($pFilename);
        // Read pictures in the Pictures Stream
        $this->loadPicturesStream();
        // Read information in the Current User Stream
        $this->loadCurrentUserStream();
        // Read information in the PowerPoint Document Stream
        $this->loadPowerpointDocumentStream();

        return $this->oPhpPresentation;
    }

    /**
     * Read OLE Part
     * @param string $pFilename
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
        $stream = $this->streamPictures;

        $pos = 0;
        do {
            $arrayRH = $this->loadRecordHeader($stream, $pos);
            $pos += 8;
            $readSuccess = false;
            if ($arrayRH['recVer'] == 0x00 && ($arrayRH['recType'] == 0xF007 || ($arrayRH['recType'] >= 0xF018 && $arrayRH['recType'] <= 0xF117))) {
                //@link : http://msdn.microsoft.com/en-us/library/dd950560(v=office.12).aspx
                if ($arrayRH['recType'] == 0xF007) {
                    // OfficeArtFBSE
                    throw new \Exception('Feature not implemented (l.'.__LINE__.')');
                }
                if ($arrayRH['recType'] >= 0xF018 && $arrayRH['recType'] <= 0xF117) {
                    $arrayRecord = $this->readRecordOfficeArtBlip($stream, $pos - 8);
                    if ($arrayRecord['length'] > 0) {
                        $pos += $arrayRecord['length'];
                        $this->arrayPictures[] = $arrayRecord['picture'];
                    }
                }
                $readSuccess = true;
            }
        } while ($readSuccess === true);
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
        do {
            $char = self::getInt1d($this->streamCurrentUser, $pos);
            if (($char >= 0x00 && $char <= 0x1F) || ($char >= 0x7F && $char <= 0x9F)) {
                $char = false;
            } else {
                $ansiUserName .= chr($char);
                $pos += 1;
            }
        } while ($char !== false);

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
        $this->readRecordUserEditAtom($this->streamPowerpointDocument, $this->offsetToCurrentEdit);

        $this->readRecordPersistDirectoryAtom($this->streamPowerpointDocument, $this->offsetPersistDirectory);

        foreach ($this->rgPersistDirEntry as $offsetDir) {
            $pos = $offsetDir;

            $rh = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
            $pos += 8;
            $this->inMainType = $rh['recType'];
            $this->currentNote = null;
            switch ($rh['recType']) {
                case self::RT_DOCUMENT:
                    $this->readRecordDocumentContainer($this->streamPowerpointDocument, $pos);
                    break;
                case self::RT_NOTES:
                    $this->readRecordNotesContainer($this->streamPowerpointDocument, $pos);
                    break;
                case self::RT_SLIDE:
                    $this->readRecordSlideContainer($this->streamPowerpointDocument, $pos);
                    break;
                default:
                    // throw new \Exception('Feature not implemented : l.'.__LINE__.'('.dechex($rh['recType']).')');
                    break;
            }
        }
    }

    /**
     * Read a record header
     * @param string $stream
     * @param integer $pos
     * @return array
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

        $ord24 = ($or24 & 127) << 24;
        if ($or24 >= 128) {
            // negative number
            $ord24 = -abs((256 - $or24) << 24);
        }
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | $ord24;
    }

    /**
     * A container record that specifies the animation and sound information for a shape.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd772900(v=office.12).aspx
     */
    private function readRecordAnimationInfoContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_ANIMATIONINFO) {
            // Record Header
            $arrayReturn['length'] += 8;
            // animationAtom
            // animationSound
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies information about the document.
     * @param string $stream
     * @param integer $pos
     * @link http://msdn.microsoft.com/en-us/library/dd947357(v=office.12).aspx
     */
    private function readRecordDocumentContainer($stream, $pos)
    {
        $documentAtom = $this->loadRecordHeader($stream, $pos);
        $pos += 8;
        if ($documentAtom['recVer'] != 0x1 || $documentAtom['recInstance'] != 0x000 || $documentAtom['recType'] != self::RT_DOCUMENTATOM) {
            throw new \Exception('File PowerPoint 97 in error (Location : RTDocument > DocumentAtom).');
        }
        $pos += $documentAtom['recLen'];

        $exObjList = $this->loadRecordHeader($stream, $pos);
        if ($exObjList['recVer'] == 0xF && $exObjList['recInstance'] == 0x000 && $exObjList['recType'] == self::RT_EXTERNALOBJECTLIST) {
            $pos += 8;
            // exObjListAtom > rh
            $exObjListAtom = $this->loadRecordHeader($stream, $pos);
            if ($exObjListAtom['recVer'] != 0x0 || $exObjListAtom['recInstance'] != 0x000 || $exObjListAtom['recType'] != self::RT_EXTERNALOBJECTLISTATOM || $exObjListAtom['recLen'] != 0x00000004) {
                throw new \Exception('File PowerPoint 97 in error (Location : RTDocument > DocumentAtom > exObjList > exObjListAtom).');
            }
            $pos += 8;
            // exObjListAtom > exObjIdSeed
            $pos += 4;
            // rgChildRec
            $exObjList['recLen'] -= 12;
            do {
                $childRec = $this->loadRecordHeader($stream, $pos);
                $pos += 8;
                $exObjList['recLen'] -= 8;
                switch ($childRec['recType']) {
                    case self::RT_EXTERNALHYPERLINK:
                        //@link : http://msdn.microsoft.com/en-us/library/dd944995(v=office.12).aspx
                        // exHyperlinkAtom > rh
                        $exHyperlinkAtom = $this->loadRecordHeader($stream, $pos);
                        if ($exHyperlinkAtom['recVer'] != 0x0 || $exHyperlinkAtom['recInstance'] != 0x000 || $exHyperlinkAtom['recType'] != self::RT_EXTERNALHYPERLINKATOM || $exObjListAtom['recLen'] != 0x00000004) {
                            throw new \Exception('File PowerPoint 97 in error (Location : RTDocument > DocumentAtom > exObjList > rgChildRec > RT_ExternalHyperlink).');
                        }
                        $pos += 8;
                        $exObjList['recLen'] -= 8;
                        // exHyperlinkAtom > exHyperlinkId
                        $exHyperlinkId = self::getInt4d($stream, $pos);
                        $pos += 4;
                        $exObjList['recLen'] -= 4;

                        $this->arrayHyperlinks[$exHyperlinkId] = array();
                        // friendlyNameAtom
                        $friendlyNameAtom  = $this->loadRecordHeader($stream, $pos);
                        if ($friendlyNameAtom['recVer'] == 0x0 && $friendlyNameAtom['recInstance'] == 0x000 && $friendlyNameAtom['recType'] == self::RT_CSTRING && $friendlyNameAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $this->arrayHyperlinks[$exHyperlinkId]['text'] = '';
                            for ($inc = 0; $inc < ($friendlyNameAtom['recLen'] / 2); $inc++) {
                                $char = self::getInt2d($stream, $pos);
                                $pos += 2;
                                $exObjList['recLen'] -= 2;
                                $this->arrayHyperlinks[$exHyperlinkId]['text'] .= chr($char);
                            }
                        }
                        // targetAtom
                        $targetAtom  = $this->loadRecordHeader($stream, $pos);
                        if ($targetAtom['recVer'] == 0x0 && $targetAtom['recInstance'] == 0x001 && $targetAtom['recType'] == self::RT_CSTRING && $targetAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $this->arrayHyperlinks[$exHyperlinkId]['url'] = '';
                            for ($inc = 0; $inc < ($targetAtom['recLen'] / 2); $inc++) {
                                $char = self::getInt2d($stream, $pos);
                                $pos += 2;
                                $exObjList['recLen'] -= 2;
                                $this->arrayHyperlinks[$exHyperlinkId]['url'] .= chr($char);
                            }
                        }
                        // locationAtom
                        $locationAtom  = $this->loadRecordHeader($stream, $pos);
                        if ($locationAtom['recVer'] == 0x0 && $locationAtom['recInstance'] == 0x003 && $locationAtom['recType'] == self::RT_CSTRING && $locationAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $string = '';
                            for ($inc = 0; $inc < ($locationAtom['recLen'] / 2); $inc++) {
                                $char = self::getInt2d($stream, $pos);
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
        $documentTextInfo = $this->loadRecordHeader($stream, $pos);
        if ($documentTextInfo['recVer'] == 0xF && $documentTextInfo['recInstance'] == 0x000 && $documentTextInfo['recType'] == self::RT_ENVIRONMENT) {
            $pos += 8;
            //@link : http://msdn.microsoft.com/en-us/library/dd952717(v=office.12).aspx
            $kinsoku = $this->loadRecordHeader($stream, $pos);
            if ($kinsoku['recVer'] == 0xF && $kinsoku['recInstance'] == 0x002 && $kinsoku['recType'] == self::RT_KINSOKU) {
                $pos += 8;
                $pos += $kinsoku['recLen'];
            }

            //@link : http://msdn.microsoft.com/en-us/library/dd948152(v=office.12).aspx
            $fontCollection = $this->loadRecordHeader($stream, $pos);
            if ($fontCollection['recVer'] == 0xF && $fontCollection['recInstance'] == 0x000 && $fontCollection['recType'] == self::RT_FONTCOLLECTION) {
                $pos += 8;
                do {
                    $fontEntityAtom = $this->loadRecordHeader($stream, $pos);
                    $pos += 8;
                    $fontCollection['recLen'] -= 8;
                    if ($fontEntityAtom['recVer'] != 0x0 || $fontEntityAtom['recInstance'] > 128 || $fontEntityAtom['recType'] != self::RT_FONTENTITYATOM) {
                        throw new \Exception('File PowerPoint 97 in error (Location : RTDocument > RT_Environment > RT_FontCollection > RT_FontEntityAtom).');
                    }
                    $string = '';
                    for ($inc = 0; $inc < 32; $inc++) {
                        $char = self::getInt2d($stream, $pos);
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

                    $fontEmbedData1 = $this->loadRecordHeader($stream, $pos);
                    if ($fontEmbedData1['recVer'] == 0x0 && $fontEmbedData1['recInstance'] >= 0x000 && $fontEmbedData1['recInstance'] <= 0x003 && $fontEmbedData1['recType'] == self::RT_FONTEMBEDDATABLOB) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData1['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData1['recLen'];
                    }

                    $fontEmbedData2 = $this->loadRecordHeader($stream, $pos);
                    if ($fontEmbedData2['recVer'] == 0x0 && $fontEmbedData2['recInstance'] >= 0x000 && $fontEmbedData2['recInstance'] <= 0x003 && $fontEmbedData2['recType'] == self::RT_FONTEMBEDDATABLOB) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData2['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData2['recLen'];
                    }

                    $fontEmbedData3 = $this->loadRecordHeader($stream, $pos);
                    if ($fontEmbedData3['recVer'] == 0x0 && $fontEmbedData3['recInstance'] >= 0x000 && $fontEmbedData3['recInstance'] <= 0x003 && $fontEmbedData3['recType'] == self::RT_FONTEMBEDDATABLOB) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData3['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData3['recLen'];
                    }

                    $fontEmbedData4 = $this->loadRecordHeader($stream, $pos);
                    if ($fontEmbedData4['recVer'] == 0x0 && $fontEmbedData4['recInstance'] >= 0x000 && $fontEmbedData4['recInstance'] <= 0x003 && $fontEmbedData4['recType'] == self::RT_FONTEMBEDDATABLOB) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData4['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData4['recLen'];
                    }
                } while ($fontCollection['recLen'] > 0);
            }

            $textCFDefaultsAtom = $this->loadRecordHeader($stream, $pos);
            if ($textCFDefaultsAtom['recVer'] == 0x0 && $textCFDefaultsAtom['recInstance'] == 0x000 && $textCFDefaultsAtom['recType'] == self::RT_TEXTCHARFORMATEXCEPTIONATOM) {
                $pos += 8;
                $pos += $textCFDefaultsAtom['recLen'];
            }

            $textPFDefaultsAtom = $this->loadRecordHeader($stream, $pos);
            if ($textPFDefaultsAtom['recVer'] == 0x0 && $textPFDefaultsAtom['recInstance'] == 0x000 && $textPFDefaultsAtom['recType'] == self::RT_TEXTPARAGRAPHFORMATEXCEPTIONATOM) {
                $pos += 8;
                $pos += $textPFDefaultsAtom['recLen'];
            }

            $defaultRulerAtom = $this->loadRecordHeader($stream, $pos);
            if ($defaultRulerAtom['recVer'] == 0x0 && $defaultRulerAtom['recInstance'] == 0x000 && $defaultRulerAtom['recType'] == self::RT_DEFAULTRULERATOM) {
                $pos += 8;
                $pos += $defaultRulerAtom['recLen'];
            }

            $textSIDefaultsAtom = $this->loadRecordHeader($stream, $pos);
            if ($textSIDefaultsAtom['recVer'] == 0x0 && $textSIDefaultsAtom['recInstance'] == 0x000 && $textSIDefaultsAtom['recType'] == self::RT_TEXTSPECIALINFODEFAULTATOM) {
                $pos += 8;
                $pos += $textSIDefaultsAtom['recLen'];
            }

            $textMasterStyleAtom = $this->loadRecordHeader($stream, $pos);
            if ($textMasterStyleAtom['recVer'] == 0x0 && $textMasterStyleAtom['recType'] == self::RT_TEXTMASTERSTYLEATOM) {
                $pos += 8;
                $pos += $textMasterStyleAtom['recLen'];
            }
        }

        $soundCollection = $this->loadRecordHeader($stream, $pos);
        if ($soundCollection['recVer'] == 0xF && $soundCollection['recInstance'] == 0x005 && $soundCollection['recType'] == self::RT_SOUNDCOLLECTION) {
            $pos += 8;
            $pos += $soundCollection['recLen'];
        }

        $drawingGroup = $this->loadRecordHeader($stream, $pos);
        if ($drawingGroup['recVer'] == 0xF && $drawingGroup['recInstance'] == 0x000 && $drawingGroup['recType'] == self::RT_DRAWINGGROUP) {
            $drawing = $this->readRecordDrawingGroupContainer($stream, $pos);
            $pos += 8;
            $pos += $drawing['length'];
        }

        $masterList = $this->loadRecordHeader($stream, $pos);
        if ($masterList['recVer'] == 0xF && $masterList['recInstance'] == 0x001 && $masterList['recType'] == self::RT_SLIDELISTWITHTEXT) {
            $pos += 8;
            $pos += $masterList['recLen'];
        }

        $docInfoList = $this->loadRecordHeader($stream, $pos);
        if ($docInfoList['recVer'] == 0xF && $docInfoList['recInstance'] == 0x000 && $docInfoList['recType'] == self::RT_LIST) {
            $pos += 8;
            $pos += $docInfoList['recLen'];
        }

        $slideHF = $this->loadRecordHeader($stream, $pos);
        if ($slideHF['recVer'] == 0xF && $slideHF['recInstance'] == 0x003 && $slideHF['recType'] == self::RT_HEADERSFOOTERS) {
            $pos += 8;
            $pos += $slideHF['recLen'];
        }

        $notesHF = $this->loadRecordHeader($stream, $pos);
        if ($notesHF['recVer'] == 0xF && $notesHF['recInstance'] == 0x004 && $notesHF['recType'] == self::RT_HEADERSFOOTERS) {
            $pos += 8;
            $pos += $notesHF['recLen'];
        }

        // SlideListWithTextContainer
        $slideList = $this->loadRecordHeader($stream, $pos);
        if ($slideList['recVer'] == 0xF && $slideList['recInstance'] == 0x000 && $slideList['recType'] == self::RT_SLIDELISTWITHTEXT) {
            $pos += 8;
            do {
                // SlideListWithTextSubContainerOrAtom
                $rhSlideList = $this->loadRecordHeader($stream, $pos);
                if ($rhSlideList['recVer'] == 0x0 && $rhSlideList['recInstance'] == 0x000 && $rhSlideList['recType'] == self::RT_SLIDEPERSISTATOM && $rhSlideList['recLen'] == 0x00000014) {
                    $pos += 8;
                    $slideList['recLen'] -= 8;
                    // persistIdRef
                    $pos += 4;
                    $slideList['recLen'] -= 4;
                    // reserved1 - fShouldCollapse - fNonOutlineData - reserved2
                    $pos += 4;
                    $slideList['recLen'] -= 4;
                    // cTexts
                    $pos += 4;
                    $slideList['recLen'] -= 4;
                    // slideId
                    $slideId = self::getInt4d($stream, $pos);
                    if ($slideId == -2147483648) {
                        $slideId = 0;
                    }
                    if ($slideId > 0) {
                        $this->arrayNotes[$this->oPhpPresentation->getActiveSlideIndex()] = $slideId;
                    }
                    $pos += 4;
                    $slideList['recLen'] -= 4;
                    // reserved3
                    $pos += 4;
                    $slideList['recLen'] -= 4;
                }
            } while ($slideList['recLen'] > 0);
        }
    }

    /**
     * An atom record that specifies information about a slide.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd923801(v=office.12).aspx
     */
    private function readRecordDrawingContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_DRAWING) {
            // Record Header
            $arrayReturn['length'] += 8;

            $officeArtDg = $this->readRecordOfficeArtDgContainer($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $officeArtDg['length'];
        }
        return $arrayReturn;
    }

    private function readRecordDrawingGroupContainer($stream, $pos)
    {

        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_DRAWINGGROUP) {
            // Record Header
            $arrayReturn['length'] += 8;
            $arrayReturn['length'] += $data['recLen'];
        }
        return $arrayReturn;
    }

    /**
     * An atom record that specifies a reference to an external object.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd910388(v=office.12).aspx
     */
    private function readRecordExObjRefAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_EXTERNALOBJECTREFATOM && $data['recLen'] == 0x00000004) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a type of action to be performed.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd953300(v=office.12).aspx
     */
    private function readRecordInteractiveInfoAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_INTERACTIVEINFOATOM && $data['recLen'] == 0x00000010) {
            // Record Header
            $arrayReturn['length'] += 8;
            // soundIdRef
            $arrayReturn['length'] += 4;
            // exHyperlinkIdRef
            $arrayReturn['exHyperlinkIdRef'] = self::getInt4d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 4;
            // action
            $arrayReturn['length'] += 1;
            // oleVerb
            $arrayReturn['length'] += 1;
            // jump
            $arrayReturn['length'] += 1;
            // fAnimated (1 bit)
            // fStopSound (1 bit)
            // fCustomShowReturn (1 bit)
            // fVisited (1 bit)
            // reserved (4 bits)
            $arrayReturn['length'] += 1;
            // hyperlinkType
            $arrayReturn['length'] += 1;
            // unused
            $arrayReturn['length'] += 3;
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies the name of a macro, a file name, or a named show.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd925121(v=office.12).aspx
     */
    private function readRecordMacroNameAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x002 && $data['recType'] == self::RT_CSTRING && $data['recLen'] % 2 == 0) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies what actions to perform when interacting with an object by means of a mouse click.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd952348(v=office.12).aspx
     */
    private function readRecordMouseClickInteractiveInfoContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_INTERACTIVEINFO) {
            // Record Header
            $arrayReturn['length'] += 8;
            // interactiveInfoAtom
            $interactiveInfoAtom = $this->readRecordInteractiveInfoAtom($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $interactiveInfoAtom['length'];
            if ($interactiveInfoAtom['length'] > 0) {
                $arrayReturn['exHyperlinkIdRef'] = $interactiveInfoAtom['exHyperlinkIdRef'];
            }
            // macroNameAtom
            $macroNameAtom = $this->readRecordMacroNameAtom($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $macroNameAtom['length'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies what actions to perform when interacting with an object by moving the mouse cursor over it.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @throws \Exception
     * @link https://msdn.microsoft.com/en-us/library/dd925811(v=office.12).aspx
     */
    private function readRecordMouseOverInteractiveInfoContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x001 && $data['recType'] == self::RT_INTERACTIVEINFO) {
            // Record Header
            $arrayReturn['length'] += 8;
            // interactiveInfoAtom
            // macroNameAtom
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtBlip record specifies BLIP file data.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @throws \Exception
     * @link https://msdn.microsoft.com/en-us/library/dd910081(v=office.12).aspx
     */
    private function readRecordOfficeArtBlip($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
            'picture' => null
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && ($data['recType'] >= 0xF018 && $data['recType'] <= 0xF117)) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            switch ($data['recType']) {
                case self::OFFICEARTBLIPJPG:
                case self::OFFICEARTBLIPPNG:
                    // rgbUid1
                    $arrayReturn['length'] += 16;
                    $data['recLen'] -= 16;
                    if ($data['recInstance'] == 0x6E1) {
                        // rgbUid2
                        $arrayReturn['length'] += 16;
                        $data['recLen'] -= 16;
                    }
                    // tag
                    $arrayReturn['length'] += 1;
                    $data['recLen'] -= 1;
                    // BLIPFileData
                    $arrayReturn['picture'] = substr($this->streamPictures, $pos + $arrayReturn['length'], $data['recLen']);
                    $arrayReturn['length'] += $data['recLen'];
                    break;
                default:
                    throw new \Exception('Feature not implemented (l.'.__LINE__.' : '.dechex($data['recType'].')'));
            }
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtChildAnchor record specifies four signed integers that specify the anchor for the shape that contains this record.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd922720(v=office.12).aspx
     */
    private function readRecordOfficeArtChildAnchor($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == 0xF00F && $data['recLen'] == 0x00000010) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['left'] = (int) self::getInt4d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 4;
            $arrayReturn['top'] = (int) self::getInt4d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 4;
            $arrayReturn['width'] = (int) self::getInt4d($stream, $pos + $arrayReturn['length']) - $arrayReturn['left'];
            $arrayReturn['length'] += 4;
            $arrayReturn['height'] = (int) self::getInt4d($stream, $pos + $arrayReturn['length']) - $arrayReturn['top'];
            $arrayReturn['length'] += 4;
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies the location of a shape.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @throws \Exception
     * @link https://msdn.microsoft.com/en-us/library/dd922797(v=office.12).aspx
     */
    private function readRecordOfficeArtClientAnchor($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == 0xF010 && ($data['recLen'] == 0x00000008 || $data['recLen'] == 0x00000010)) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            switch ($data['recLen']) {
                case 0x00000008:
                    $arrayReturn['top'] = (int) (self::getInt2d($stream, $pos + $arrayReturn['length']) / 6);
                    $arrayReturn['length'] += 2;
                    $arrayReturn['left'] = (int) (self::getInt2d($stream, $pos + $arrayReturn['length']) / 6);
                    $arrayReturn['length'] += 2;
                    $arrayReturn['width'] = (int) (self::getInt2d($stream, $pos + $arrayReturn['length']) / 6) - $arrayReturn['left'];
                    $arrayReturn['length'] += 2;
                    $arrayReturn['height'] = (int) (self::getInt2d($stream, $pos + $arrayReturn['length']) / 6) - $arrayReturn['left'];
                    $arrayReturn['length'] += 2;
                    $pos += 8;
                    break;
                case 0x00000010:
                    throw new \Exception('PowerPoint97 Reader : record OfficeArtClientAnchor (0x00000010)');
            }
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies text related data for a shape.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd910958(v=office.12).aspx
     */
    private function readRecordOfficeArtClientTextbox($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
            'text' => '',
            'numParts' => 0,
            'numTexts' => 0,
            'hyperlink' => array(),
        );

        $data = $this->loadRecordHeader($stream, $pos);
        // recVer 0xF
        // Doc : 0x0    https://msdn.microsoft.com/en-us/library/dd910958(v=office.12).aspx
        // Sample : 0xF https://msdn.microsoft.com/en-us/library/dd953497(v=office.12).aspx
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == 0xF00D) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $strLen = 0;
            do {
                $rhChild = $this->loadRecordHeader($stream, $pos + $arrayReturn['length']);
                /**
                 * @link : https://msdn.microsoft.com/en-us/library/dd947039(v=office.12).aspx
                 */
                // echo dechex($rhChild['recType']).'-'.$rhChild['recType'].EOL;
                switch ($rhChild['recType']) {
                    case self::RT_INTERACTIVEINFO:
                        //@link : http://msdn.microsoft.com/en-us/library/dd948623(v=office.12).aspx
                        if ($rhChild['recInstance'] == 0x0000) {
                            $mouseClickInfo = $this->readRecordMouseClickInteractiveInfoContainer($stream, $pos + $arrayReturn['length']);
                            $arrayReturn['length'] += $mouseClickInfo['length'];
                            $arrayReturn['hyperlink'][]['id'] = $mouseClickInfo['exHyperlinkIdRef'];
                        }
                        if ($rhChild['recInstance'] == 0x0001) {
                            $mouseOverInfo = $this->readRecordMouseOverInteractiveInfoContainer($stream, $pos + $arrayReturn['length']);
                            $arrayReturn['length'] += $mouseOverInfo['length'];
                        }
                        break;
                    case self::RT_STYLETEXTPROPATOM:
                        $arrayReturn['length'] += 8;
                        // @link : http://msdn.microsoft.com/en-us/library/dd950647(v=office.12).aspx
                        // rgTextPFRun
                        $strLenRT = $strLen + 1;
                        do {
                            $strucTextPFRun = $this->readStructureTextPFRun($stream, $pos + $arrayReturn['length'], $strLenRT);
                            $arrayReturn['numTexts']++;
                            $arrayReturn['text'.$arrayReturn['numTexts']] = $strucTextPFRun;
                            if (isset($strucTextPFRun['alignH'])) {
                                $arrayReturn['alignH'] = $strucTextPFRun['alignH'];
                            }
                            $strLenRT = $strucTextPFRun['strLenRT'];
                            $arrayReturn['length'] += $strucTextPFRun['length'];
                        } while ($strLenRT > 0);
                        // rgTextCFRun
                        $strLenRT = $strLen + 1;
                        do {
                            $strucTextCFRun = $this->readStructureTextCFRun($stream, $pos + $arrayReturn['length'], $strLenRT);
                            $arrayReturn['numParts']++;
                            $arrayReturn['part'.$arrayReturn['numParts']] = $strucTextCFRun;
                            $strLenRT = $strucTextCFRun['strLenRT'];
                            $arrayReturn['length'] += $strucTextCFRun['length'];
                        } while ($strLenRT > 0);
                        break;
                    case self::RT_TEXTBYTESATOM:
                        $arrayReturn['length'] += 8;
                        // @link : https://msdn.microsoft.com/en-us/library/dd947905(v=office.12).aspx
                        $strLen = (int)$rhChild['recLen'];
                        for ($inc = 0; $inc < $strLen; $inc++) {
                            $char = self::getInt1d($stream, $pos + $arrayReturn['length']);
                            if ($char == 0x0B) {
                                $char = 0x20;
                            }
                            $arrayReturn['text'] .= Text::chr($char);
                            $arrayReturn['length'] += 1;
                        }
                        break;
                    case self::RT_TEXTCHARSATOM:
                        $arrayReturn['length'] += 8;
                        // @link : http://msdn.microsoft.com/en-us/library/dd772921(v=office.12).aspx
                        $strLen = (int)($rhChild['recLen']/2);
                        for ($inc = 0; $inc < $strLen; $inc++) {
                            $char = self::getInt2d($stream, $pos + $arrayReturn['length']);
                            if ($char == 0x0B) {
                                $char = 0x20;
                            }
                            $arrayReturn['text'] .= Text::chr($char);
                            $arrayReturn['length'] += 2;
                        }
                        break;
                    case self::RT_TEXTHEADERATOM:
                        $arrayReturn['length'] += 8;
                        // @link : http://msdn.microsoft.com/en-us/library/dd905272(v=office.12).aspx
                        // textType
                        $arrayReturn['length'] += 4;
                        break;
                    case self::RT_TEXTINTERACTIVEINFOATOM:
                        $arrayReturn['length'] += 8;
                        //@link : http://msdn.microsoft.com/en-us/library/dd947973(v=office.12).aspx
                        if ($rhChild['recInstance'] == 0x0000) {
                            //@todo : MouseClickTextInteractiveInfoAtom
                            $arrayReturn['hyperlink'][count($arrayReturn['hyperlink']) - 1]['start'] = self::getInt4d($stream, $pos +  + $arrayReturn['length']);
                            $arrayReturn['length'] += 4;

                            $arrayReturn['hyperlink'][count($arrayReturn['hyperlink']) - 1]['end'] = self::getInt4d($stream, $pos +  + $arrayReturn['length']);
                            $arrayReturn['length'] += 4;
                        }
                        if ($rhChild['recInstance'] == 0x0001) {
                            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
                        }
                        break;
                    case self::RT_TEXTSPECIALINFOATOM:
                        $arrayReturn['length'] += 8;
                        // @link : http://msdn.microsoft.com/en-us/library/dd945296(v=office.12).aspx
                        $strLenRT = $strLen + 1;
                        do {
                            $structTextSIRun = $this->readStructureTextSIRun($stream, $pos + $arrayReturn['length'], $strLenRT);
                            $strLenRT = $structTextSIRun['strLenRT'];
                            $arrayReturn['length'] += $structTextSIRun['length'];
                        } while ($strLenRT > 0);
                        break;
                    case self::RT_TEXTRULERATOM:
                        $arrayReturn['length'] += 8;
                        // @link : http://msdn.microsoft.com/en-us/library/dd953212(v=office.12).aspx
                        $structRuler = $this->readStructureTextRuler($stream, $pos + $arrayReturn['length']);
                        $arrayReturn['length'] += $structRuler['length'];
                        break;
                    case self::RT_SLIDENUMBERMETACHARATOM:
                        $datasRecord = $this->readRecordSlideNumberMCAtom($stream, $pos + $arrayReturn['length']);
                        $arrayReturn['length'] += $datasRecord['length'];
                        break;
                    default:
                        $arrayReturn['length'] += 8;
                        $arrayReturn['length'] += $rhChild['recLen'];
                    // throw new \Exception('Feature not implemented (l.'.__LINE__.' : 0x'.dechex($rhChild['recType']).')');
                }
            } while (($data['recLen'] - $arrayReturn['length']) > 0);
        }
        return $arrayReturn;
    }

    /**
     * The OfficeArtSpContainer record specifies a shape container.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd943794(v=office.12).aspx
     */
    private function readRecordOfficeArtSpContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
            'shape' => null,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == 0xF004) {
            // Record Header
            $arrayReturn['length'] += 8;
            // shapeGroup
            $shapeGroup = $this->readRecordOfficeArtFSPGR($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $shapeGroup['length'];

            // shapeProp
            $shapeProp = $this->readRecordOfficeArtFSP($stream, $pos + $arrayReturn['length']);
            if ($shapeProp['length'] == 0) {
                throw new \Exception('PowerPoint97 Reader : record OfficeArtFSP');
            }
            $arrayReturn['length'] += $shapeProp['length'];

            if ($shapeProp['fDeleted'] == 0x1 && $shapeProp['fChild'] == 0x0) {
                // deletedShape
                $deletedShape = $this->readRecordOfficeArtFPSPL($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += $deletedShape['length'];
            }

            // shapePrimaryOptions
            $shpPrimaryOptions = $this->readRecordOfficeArtFOPT($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $shpPrimaryOptions['length'];

            // shapeSecondaryOptions1
            $shpSecondaryOptions1 = $this->readRecordOfficeArtSecondaryFOPT($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $shpSecondaryOptions1['length'];

            // shapeTertiaryOptions1
            $shpTertiaryOptions1 = $this->readRecordOfficeArtTertiaryFOPT($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $shpTertiaryOptions1['length'];

            // childAnchor
            $childAnchor = $this->readRecordOfficeArtChildAnchor($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $childAnchor['length'];

            // clientAnchor
            $clientAnchor = $this->readRecordOfficeArtClientAnchor($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $clientAnchor['length'];

            // clientData
            $clientData = $this->readRecordOfficeArtClientData($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $clientData['length'];

            // clientTextbox
            $clientTextbox = $this->readRecordOfficeArtClientTextbox($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $clientTextbox['length'];

            // shapeSecondaryOptions2
            if ($shpSecondaryOptions1['length'] == 0) {
                $shpSecondaryOptions2 = $this->readRecordOfficeArtSecondaryFOPT($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += $shpSecondaryOptions2['length'];
            }

            // shapeTertiaryOptions2
            if ($shpTertiaryOptions1['length'] == 0) {
                $shpTertiaryOptions2 = $this->readRecordOfficeArtTertiaryFOPT($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += $shpTertiaryOptions2['length'];
            }

            // Core : Shape
            // Informations about group are not defined
            $arrayDimensions = array();
            $bIsGroup = false;
            if (is_object($this->oCurrentGroup)) {
                if (!$this->bFirstShapeGroup) {
                    if ($clientAnchor['length'] > 0) {
                        // $this->oCurrentGroup->setOffsetX($clientAnchor['left']);
                        // $this->oCurrentGroup->setOffsetY($clientAnchor['top']);
                        // $this->oCurrentGroup->setHeight($clientAnchor['height']);
                        // $this->oCurrentGroup->setWidth($clientAnchor['width']);
                    }
                    $bIsGroup = true;
                    $this->bFirstShapeGroup = true;
                } else {
                    if ($childAnchor['length'] > 0) {
                        $arrayDimensions = $childAnchor;
                    }
                }
            } else {
                if ($clientAnchor['length'] > 0) {
                    $arrayDimensions = $clientAnchor;
                }
            }
            if (!$bIsGroup) {
                // *** Shape ***
                if (isset($shpPrimaryOptions['pib'])) {
                    // isDrawing
                    $drawingPib = $shpPrimaryOptions['pib'];
                    if (isset($this->arrayPictures[$drawingPib - 1])) {
                        $gdImage = imagecreatefromstring($this->arrayPictures[$drawingPib - 1]);
                        $arrayReturn['shape'] = new Drawing\Gd();
                        $arrayReturn['shape']->setImageResource($gdImage);
                    }
                } elseif (isset($shpPrimaryOptions['line']) && $shpPrimaryOptions['line']) {
                    // isLine
                    $arrayReturn['shape'] = new Line(0, 0, 0, 0);
                } elseif ($clientTextbox['length'] > 0) {
                    $arrayReturn['shape'] = new RichText();
                    if (isset($clientTextbox['alignH'])) {
                        $arrayReturn['shape']->getActiveParagraph()->getAlignment()->setHorizontal($clientTextbox['alignH']);
                    }

                    $start = 0;
                    $lastLevel = -1;
                    $lastMarginLeft = 0;
                    for ($inc = 1; $inc <= $clientTextbox['numParts']; $inc++) {
                        if ($clientTextbox['numParts'] == $clientTextbox['numTexts'] && isset($clientTextbox['text'.$inc])) {
                            if (isset($clientTextbox['text'.$inc]['bulletChar'])) {
                                $arrayReturn['shape']->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
                                $arrayReturn['shape']->getActiveParagraph()->getBulletStyle()->setBulletChar($clientTextbox['text'.$inc]['bulletChar']);
                            }
                            // Indent
                            $indent = 0;
                            if (isset($clientTextbox['text'.$inc]['indent'])) {
                                $indent = $clientTextbox['text'.$inc]['indent'];
                            }
                            if (isset($clientTextbox['text'.$inc]['leftMargin'])) {
                                if ($lastMarginLeft > $clientTextbox['text'.$inc]['leftMargin']) {
                                    $lastLevel--;
                                }
                                if ($lastMarginLeft < $clientTextbox['text'.$inc]['leftMargin']) {
                                    $lastLevel++;
                                }
                                $arrayReturn['shape']->getActiveParagraph()->getAlignment()->setLevel($lastLevel);
                                $lastMarginLeft = $clientTextbox['text'.$inc]['leftMargin'];

                                $arrayReturn['shape']->getActiveParagraph()->getAlignment()->setMarginLeft($clientTextbox['text'.$inc]['leftMargin']);
                                $arrayReturn['shape']->getActiveParagraph()->getAlignment()->setIndent($indent - $clientTextbox['text'.$inc]['leftMargin']);
                            }
                        }
                        // Texte
                        $sText = substr(isset($clientTextbox['text']) ? $clientTextbox['text'] : '', $start, $clientTextbox['part'.$inc]['partLength']);
                        $sHyperlinkURL = '';
                        if (empty($sText)) {
                            // Is there a hyperlink ?
                            if (isset($clientTextbox['hyperlink']) && is_array($clientTextbox['hyperlink']) && !empty($clientTextbox['hyperlink'])) {
                                foreach ($clientTextbox['hyperlink'] as $itmHyperlink) {
                                    if ($itmHyperlink['start'] == $start && ($itmHyperlink['end'] - $itmHyperlink['start']) == $clientTextbox['part'.$inc]['partLength']) {
                                        $sText = $this->arrayHyperlinks[$itmHyperlink['id']]['text'];
                                        $sHyperlinkURL = $this->arrayHyperlinks[$itmHyperlink['id']]['url'];
                                        break;
                                    }
                                }
                            }
                        }
                        // New paragraph
                        $bCreateParagraph = false;
                        if (strpos($sText, "\r") !== false) {
                            $bCreateParagraph = true;
                            $sText = str_replace("\r", '', $sText);
                        }
                        // TextRun
                        $txtRun = $arrayReturn['shape']->createTextRun($sText);
                        if (isset($clientTextbox['part'.$inc]['bold'])) {
                            $txtRun->getFont()->setBold($clientTextbox['part'.$inc]['bold']);
                        }
                        if (isset($clientTextbox['part'.$inc]['italic'])) {
                            $txtRun->getFont()->setItalic($clientTextbox['part'.$inc]['italic']);
                        }
                        if (isset($clientTextbox['part'.$inc]['underline'])) {
                            $txtRun->getFont()->setUnderline($clientTextbox['part'.$inc]['underline']);
                        }
                        if (isset($clientTextbox['part'.$inc]['fontName'])) {
                            $txtRun->getFont()->setName($clientTextbox['part'.$inc]['fontName']);
                        }
                        if (isset($clientTextbox['part'.$inc]['fontSize'])) {
                            $txtRun->getFont()->setSize($clientTextbox['part'.$inc]['fontSize']);
                        }
                        if (isset($clientTextbox['part'.$inc]['color'])) {
                            $txtRun->getFont()->setColor($clientTextbox['part'.$inc]['color']);
                        }
                        // Hyperlink
                        if (!empty($sHyperlinkURL)) {
                            $txtRun->setHyperlink(new Hyperlink($sHyperlinkURL));
                        }

                        $start += $clientTextbox['part'.$inc]['partLength'];
                        if ($bCreateParagraph) {
                            $arrayReturn['shape']->createParagraph();
                        }
                    }
                }

                // *** Properties ***
                // Dimensions
                if ($arrayReturn['shape'] instanceof AbstractShape) {
                    if (!empty($arrayDimensions)) {
                        $arrayReturn['shape']->setOffsetX($arrayDimensions['left']);
                        $arrayReturn['shape']->setOffsetY($arrayDimensions['top']);
                        $arrayReturn['shape']->setHeight($arrayDimensions['height']);
                        $arrayReturn['shape']->setWidth($arrayDimensions['width']);
                    }
                    // Rotation
                    if (isset($shpPrimaryOptions['rotation'])) {
                        $rotation = $shpPrimaryOptions['rotation'];
                        $arrayReturn['shape']->setRotation($rotation);
                    }
                    // Shadow
                    if (isset($shpPrimaryOptions['shadowOffsetX']) && isset($shpPrimaryOptions['shadowOffsetY'])) {
                        $shadowOffsetX = $shpPrimaryOptions['shadowOffsetX'];
                        $shadowOffsetY = $shpPrimaryOptions['shadowOffsetY'];
                        if ($shadowOffsetX != 0 && $shadowOffsetX != 0) {
                            $arrayReturn['shape']->getShadow()->setVisible(true);
                            if ($shadowOffsetX > 0 && $shadowOffsetX == $shadowOffsetY) {
                                $arrayReturn['shape']->getShadow()->setDistance($shadowOffsetX)->setDirection(45);
                            }
                        }
                    }
                    // Specific Line
                    if ($arrayReturn['shape'] instanceof Line) {
                        if (isset($shpPrimaryOptions['lineColor'])) {
                            $arrayReturn['shape']->getBorder()->getColor()->setARGB('FF'.$shpPrimaryOptions['lineColor']);
                        }
                        if (isset($shpPrimaryOptions['lineWidth'])) {
                            $arrayReturn['shape']->setHeight($shpPrimaryOptions['lineWidth']);
                        }
                    }
                    // Specific RichText
                    if ($arrayReturn['shape'] instanceof RichText) {
                        if (isset($shpPrimaryOptions['insetBottom'])) {
                            $arrayReturn['shape']->setInsetBottom($shpPrimaryOptions['insetBottom']);
                        }
                        if (isset($shpPrimaryOptions['insetLeft'])) {
                            $arrayReturn['shape']->setInsetLeft($shpPrimaryOptions['insetLeft']);
                        }
                        if (isset($shpPrimaryOptions['insetRight'])) {
                            $arrayReturn['shape']->setInsetRight($shpPrimaryOptions['insetRight']);
                        }
                        if (isset($shpPrimaryOptions['insetTop'])) {
                            $arrayReturn['shape']->setInsetTop($shpPrimaryOptions['insetTop']);
                        }
                    }
                }
            } else {
                // Rotation
                if (isset($shpPrimaryOptions['rotation'])) {
                    $rotation = $shpPrimaryOptions['rotation'];
                    $this->oCurrentGroup->setRotation($rotation);
                }
            }
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtSpgrContainer record specifies a container for groups of shapes.
     * @param string $stream
     * @param integer $pos
     * @param boolean $bInGroup
     * @return array
     * @link : https://msdn.microsoft.com/en-us/library/dd910416(v=office.12).aspx
     */
    private function readRecordOfficeArtSpgrContainer($stream, $pos, $bInGroup = false)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == 0xF003) {
            $arrayReturn['length'] += 8;

            do {
                $rhFileBlock = $this->loadRecordHeader($stream, $pos + $arrayReturn['length']);
                if (!($rhFileBlock['recVer'] == 0xF && $rhFileBlock['recInstance'] == 0x0000 && ($rhFileBlock['recType'] == 0xF003 || $rhFileBlock['recType'] == 0xF004))) {
                    throw new \Exception('PowerPoint97 Reader : readRecordOfficeArtSpgrContainer.');
                }

                switch ($rhFileBlock['recType']) {
                    case 0xF003:
                        // Core
                        $this->oCurrentGroup = $this->oPhpPresentation->getActiveSlide()->createGroup();
                        $this->bFirstShapeGroup = false;
                        // OfficeArtSpgrContainer
                        $fileBlock = $this->readRecordOfficeArtSpgrContainer($stream, $pos + $arrayReturn['length'], true);
                        $arrayReturn['length'] += $fileBlock['length'];
                        $data['recLen'] -= $fileBlock['length'];
                        break;
                    case 0xF004:
                        // Core
                        if (!$bInGroup) {
                            $this->oCurrentGroup = null;
                        }
                        // OfficeArtSpContainer
                        $fileBlock = $this->readRecordOfficeArtSpContainer($stream, $pos + $arrayReturn['length']);
                        $arrayReturn['length'] += $fileBlock['length'];
                        $data['recLen'] -= $fileBlock['length'];
                        // Core
                        //@todo
                        if (!is_null($fileBlock['shape'])) {
                            switch ($this->inMainType) {
                                case self::RT_NOTES:
                                    $arrayIdxSlide = array_flip($this->arrayNotes);
                                    if ($this->currentNote > 0 && isset($arrayIdxSlide[$this->currentNote])) {
                                        $oSlide = $this->oPhpPresentation->getSlide($arrayIdxSlide[$this->currentNote]);
                                        if ($oSlide->getNote()->getShapeCollection()->count() == 0) {
                                            $oSlide->getNote()->addShape($fileBlock['shape']);
                                        }
                                    }
                                    break;
                                case self::RT_SLIDE:
                                    if ($bInGroup) {
                                        $this->oCurrentGroup->addShape($fileBlock['shape']);
                                    } else {
                                        $this->oPhpPresentation->getActiveSlide()->addShape($fileBlock['shape']);
                                    }
                                    break;
                            }
                        }

                        break;
                }
            } while ($data['recLen'] > 0);
        }
        return $arrayReturn;
    }

    /**
     * The OfficeArtTertiaryFOPT record specifies a table of OfficeArtRGFOPTE records,.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd950206(v=office.12).aspx
     */
    private function readRecordOfficeArtTertiaryFOPT($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x3 && $data['recType'] == 0xF122) {
            // Record Header
            $arrayReturn['length'] += 8;

            $officeArtFOPTE = array();
            for ($inc = 0; $inc < $data['recInstance']; $inc++) {
                $opid = self::getInt2d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $optOp = self::getInt4d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 4;
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
                    case 0x039F:
                        // Table properties
                        //@link : https://msdn.microsoft.com/en-us/library/dd922773(v=office.12).aspx
                        break;
                    case 0x03A0:
                        // Table Row Properties
                        //@link : https://msdn.microsoft.com/en-us/library/dd923419(v=office.12).aspx
                        if ($opt['fComplex'] == 0x1) {
                            $arrayReturn['length'] += $opt['op'];
                        }
                        break;
                    case 0x03A9:
                        // GroupShape : metroBlob
                        //@link : https://msdn.microsoft.com/en-us/library/dd943388(v=office.12).aspx
                        if ($opt['fComplex'] == 0x1) {
                            $arrayReturn['length'] += $opt['op'];
                        }
                        break;
                    case 0x01FF:
                        // Line Style Boolean
                        //@link : https://msdn.microsoft.com/en-us/library/dd951605(v=office.12).aspx
                        break;
                    default:
                        throw new \Exception('Feature not implemented (l.'.__LINE__.' : 0x'.dechex($opt['opid']).')');
                }
            }
        }
        return $arrayReturn;
    }

    /**
     * The OfficeArtDgContainer record specifies the container for all the file records for the objects in a drawing.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link : https://msdn.microsoft.com/en-us/library/dd924455(v=office.12).aspx
     */
    private function readRecordOfficeArtDgContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == 0xF002) {
            // Record Header
            $arrayReturn['length'] += 8;
            // drawingData
            $drawingData  = $this->readRecordOfficeArtFDG($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $drawingData['length'];
            // regroupItems
            //@todo
            // groupShape
            $groupShape = $this->readRecordOfficeArtSpgrContainer($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $groupShape['length'];
            // shape
            $shape = $this->readRecordOfficeArtSpContainer($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $shape['length'];
            // solvers1
            //@todo
            // deletedShapes
            //@todo
            // solvers1
            //@todo
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtFDG record specifies the number of shapes, the drawing identifier, and the shape identifier of the last shape in a drawing.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link : https://msdn.microsoft.com/en-us/library/dd946757(v=office.12).aspx
     */
    private function readRecordOfficeArtFDG($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] <= 0xFFE && $data['recType'] == 0xF008 && $data['recLen'] == 0x00000008) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtFOPT record specifies a table of OfficeArtRGFOPTE records.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd943404(v=office.12).aspx
     */
    private function readRecordOfficeArtFOPT($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x3 && $data['recType'] == 0xF00B) {
            // Record Header
            $arrayReturn['length'] += 8;

            //@link : http://msdn.microsoft.com/en-us/library/dd906086(v=office.12).aspx
            $officeArtFOPTE = array();
            for ($inc = 0; $inc < $data['recInstance']; $inc++) {
                $opid = self::getInt2d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $data['recLen'] -= 2;
                $optOp = self::getInt4d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 4;
                $data['recLen'] -= 4;
                $officeArtFOPTE[] = array(
                    'opid' => ($opid >> 0) & bindec('11111111111111'),
                    'fBid' => ($opid >> 14) & bindec('1'),
                    'fComplex' => ($opid >> 15) & bindec('1'),
                    'op' => $optOp,
                );
            }
            //@link : http://code.metager.de/source/xref/kde/calligra/filters/libmso/OPID
            foreach ($officeArtFOPTE as $opt) {
                // echo $opt['opid'].'-0x'.dechex($opt['opid']).EOL;
                switch ($opt['opid']) {
                    case 0x0004:
                        // Transform : rotation
                        //@link : https://msdn.microsoft.com/en-us/library/dd949750(v=office.12).aspx
                        $arrayReturn['rotation'] = $opt['op'];
                        break;
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
                        $arrayReturn['insetLeft'] = \PhpOffice\Common\Drawing::emuToPixels($opt['op']);
                        break;
                    case 0x0082:
                        // Text : dyTextTop
                        //@link : http://msdn.microsoft.com/en-us/library/dd925068(v=office.12).aspx
                        $arrayReturn['insetTop'] = \PhpOffice\Common\Drawing::emuToPixels($opt['op']);
                        break;
                    case 0x0083:
                        // Text : dxTextRight
                        //@link : http://msdn.microsoft.com/en-us/library/dd906782(v=office.12).aspx
                        $arrayReturn['insetRight'] = \PhpOffice\Common\Drawing::emuToPixels($opt['op']);
                        break;
                    case 0x0084:
                        // Text : dyTextBottom
                        //@link : http://msdn.microsoft.com/en-us/library/dd772858(v=office.12).aspx
                        $arrayReturn['insetBottom'] = \PhpOffice\Common\Drawing::emuToPixels($opt['op']);
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
                            $arrayReturn['pib'] = $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        } else {
                            // pib Complex
                        }
                        break;
                    case 0x13F:
                        // Blip Boolean Properties
                        //@link : https://msdn.microsoft.com/en-us/library/dd944215(v=office.12).aspx
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
                        $arrayReturn['line'] = true;
                        break;
                    case 0x145:
                        // Geometry : pVertices
                        //@link : http://msdn.microsoft.com/en-us/library/dd949814(v=office.12).aspx
                        if ($opt['fComplex'] == 1) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }
                        break;
                    case 0x146:
                        // Geometry : pSegmentInfo
                        //@link : http://msdn.microsoft.com/en-us/library/dd905742(v=office.12).aspx
                        if ($opt['fComplex'] == 1) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }
                        break;
                    case 0x155:
                        // Geometry : pAdjustHandles
                        //@link : http://msdn.microsoft.com/en-us/library/dd905890(v=office.12).aspx
                        if ($opt['fComplex'] == 1) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }
                        break;
                    case 0x156:
                        // Geometry : pGuides
                        //@link : http://msdn.microsoft.com/en-us/library/dd910801(v=office.12).aspx
                        if ($opt['fComplex'] == 1) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }
                        break;
                    case 0x157:
                        // Geometry : pInscribe
                        //@link : http://msdn.microsoft.com/en-us/library/dd904889(v=office.12).aspx
                        if ($opt['fComplex'] == 1) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
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
                        $strColor = str_pad(dechex(($opt['op'] >> 0) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        $strColor .= str_pad(dechex(($opt['op'] >> 8) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        $strColor .= str_pad(dechex(($opt['op'] >> 16) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        // echo 'fillColor  : '.$strColor.EOL;
                        break;
                    case 0x0183:
                        // Fill : fillBackColor
                        //@link : http://msdn.microsoft.com/en-us/library/dd950634(v=office.12).aspx
                        $strColor = str_pad(dechex(($opt['op'] >> 0) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        $strColor .= str_pad(dechex(($opt['op'] >> 8) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        $strColor .= str_pad(dechex(($opt['op'] >> 16) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        // echo 'fillBackColor  : '.$strColor.EOL;
                        break;
                    case 0x0193:
                        // Fill : fillRectRight
                        //@link : http://msdn.microsoft.com/en-us/library/dd951294(v=office.12).aspx
                        // echo 'fillRectRight  : '.\PhpOffice\Common\Drawing::emuToPixels($opt['op']).EOL;
                        break;
                    case 0x0194:
                        // Fill : fillRectBottom
                        //@link : http://msdn.microsoft.com/en-us/library/dd910194(v=office.12).aspx
                        // echo 'fillRectBottom   : '.\PhpOffice\Common\Drawing::emuToPixels($opt['op']).EOL;
                        break;
                    case 0x01BF:
                        // Fill : Fill Style Boolean Properties
                        //@link : http://msdn.microsoft.com/en-us/library/dd909380(v=office.12).aspx
                        break;
                    case 0x01C0:
                        // Line Style : lineColor
                        //@link : http://msdn.microsoft.com/en-us/library/dd920397(v=office.12).aspx
                        $strColor = str_pad(dechex(($opt['op'] >> 0) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        $strColor .= str_pad(dechex(($opt['op'] >> 8) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        $strColor .= str_pad(dechex(($opt['op'] >> 16) & bindec('11111111')), 2, STR_PAD_LEFT, '0');
                        $arrayReturn['lineColor'] = $strColor;
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
                        $arrayReturn['lineWidth'] = \PhpOffice\Common\Drawing::emuToPixels($opt['op']);
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
                        $arrayReturn['shadowOffsetX'] = \PhpOffice\Common\Drawing::emuToPixels($opt['op']);
                        break;
                    case 0x0206:
                        // Shadow Style : shadowOffsetY
                        //@link : http://msdn.microsoft.com/en-us/library/dd907855(v=office.12).aspx
                        $arrayReturn['shadowOffsetY'] = \PhpOffice\Common\Drawing::emuToPixels($opt['op']);
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
                    case 0x0380:
                        // Group Shape Property Set : wzName
                        //@link : http://msdn.microsoft.com/en-us/library/dd950681(v=office.12).aspx
                        if ($opt['fComplex'] == 1) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }
                        break;
                    case 0x03BF:
                        // Group Shape Property Set : Group Shape Boolean Properties
                        //@link : http://msdn.microsoft.com/en-us/library/dd949807(v=office.12).aspx
                        break;
                    default:
                        // throw new \Exception('Feature not implemented (l.'.__LINE__.' : 0x'.dechex($opt['opid']).')');
                }
            }
            if ($data['recLen'] > 0) {
                $arrayReturn['length'] += $data['recLen'];
            }
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtFPSPL record specifies the former hierarchical position of the containing object that is either a shape or a group of shapes.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd947479(v=office.12).aspx
     */
    private function readRecordOfficeArtFPSPL($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == 0xF11D && $data['recLen'] == 0x00000004) {
            $arrayReturn['length'] += 8;
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtFSP record specifies an instance of a shape.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd925898(v=office.12).aspx
     */
    private function readRecordOfficeArtFSP($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x2 && $data['recType'] == 0xF00A && $data['recLen'] == 0x00000008) {
            $arrayReturn['length'] += 8;
            // spid
            $arrayReturn['length'] += 4;
            // data
            $data = self::getInt4d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 4;
            $arrayReturn['fGroup'] = ($data >> 0) & bindec('1');
            $arrayReturn['fChild'] = ($data >> 1) & bindec('1');
            $arrayReturn['fPatriarch'] = ($data >> 2) & bindec('1');
            $arrayReturn['fDeleted'] = ($data >> 3) & bindec('1');
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtFSPGR record specifies the coordinate system of the group shape that the anchors of the child shape are expressed in.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd925381(v=office.12).aspx
     */
    private function readRecordOfficeArtFSPGR($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x1 && $data['recInstance'] == 0x000 && $data['recType'] == 0xF009 && $data['recLen'] == 0x00000010) {
            $arrayReturn['length'] += 8;
            //$arrShapeGroup['xLeft'] = self::getInt4d($this->streamPowerpointDocument, $pos);
            $arrayReturn['length'] += 4;
            //$arrShapeGroup['yTop'] = self::getInt4d($this->streamPowerpointDocument, $pos);
            $arrayReturn['length'] += 4;
            //$arrShapeGroup['xRight'] = self::getInt4d($this->streamPowerpointDocument, $pos);
            $arrayReturn['length'] += 4;
            //$arrShapeGroup['yBottom'] = self::getInt4d($this->streamPowerpointDocument, $pos);
            $arrayReturn['length'] += 4;
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtSecondaryFOPT record specifies a table of OfficeArtRGFOPTE records.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd950259(v=office.12).aspx
     */
    private function readRecordOfficeArtSecondaryFOPT($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x3 && $data['recType'] == 0xF121) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }
        return $arrayReturn;
    }

    /**
     * A container record that specifies information about a shape.
     * @param string $stream
     * @param integer $pos
     * @link : https://msdn.microsoft.com/en-us/library/dd950927(v=office.12).aspx
     */
    private function readRecordOfficeArtClientData($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == 0xF011) {
            $arrayReturn['length'] += 8;
            // shapeFlagsAtom (9 bytes)
            $dataShapeFlagsAtom = $this->readRecordShapeFlagsAtom($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $dataShapeFlagsAtom['length'];

            // shapeFlags10Atom (9 bytes)
            $dataShapeFlags10Atom = $this->readRecordShapeFlags10Atom($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $dataShapeFlags10Atom['length'];

            // exObjRefAtom (12 bytes)
            $dataExObjRefAtom = $this->readRecordExObjRefAtom($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $dataExObjRefAtom['length'];

            // animationInfo (variable)
            $dataAnimationInfo = $this->readRecordAnimationInfoContainer($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $dataAnimationInfo['length'];

            // mouseClickInteractiveInfo (variable)
            $mouseClickInfo = $this->readRecordMouseClickInteractiveInfoContainer($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $mouseClickInfo['length'];

            // mouseOverInteractiveInfo (variable)
            $mouseOverInfo = $this->readRecordMouseOverInteractiveInfoContainer($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $mouseOverInfo['length'];

            // placeholderAtom (16 bytes)
            $dataPlaceholderAtom = $this->readRecordPlaceholderAtom($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $dataPlaceholderAtom['length'];

            // recolorInfoAtom (variable)
            $dataRecolorInfo = $this->readRecordRecolorInfoAtom($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $dataRecolorInfo['length'];

            // rgShapeClientRoundtripData (variable)
            $array = array(
                self::RT_PROGTAGS,
                self::RT_ROUNDTRIPNEWPLACEHOLDERID12ATOM,
                self::RT_ROUNDTRIPSHAPEID12ATOM,
                self::RT_ROUNDTRIPHFPLACEHOLDER12ATOM,
                self::RT_ROUNDTRIPSHAPECHECKSUMFORCL12ATOM,
            );
            do {
                $dataHeaderRG = $this->loadRecordHeader($stream, $pos + $arrayReturn['length']);
                if (in_array($dataHeaderRG['recType'], $array)) {
                    switch ($dataHeaderRG['recType']) {
                        case self::RT_PROGTAGS:
                            $dataRG = $this->readRecordShapeProgTagsContainer($stream, $pos + $arrayReturn['length']);
                            $arrayReturn['length'] += $dataRG['length'];
                            break;
                        case self::RT_ROUNDTRIPHFPLACEHOLDER12ATOM:
                            $dataRG = $this->readRecordRoundTripHFPlaceholder12Atom($stream, $pos + $arrayReturn['length']);
                            $arrayReturn['length'] += $dataRG['length'];
                            break;
                        case self::RT_ROUNDTRIPSHAPEID12ATOM:
                            $dataRG = $this->readRecordRoundTripShapeId12Atom($stream, $pos + $arrayReturn['length']);
                            $arrayReturn['length'] += $dataRG['length'];
                            break;
                        default:
                            throw new \Exception('Feature not implemented (l.'.__LINE__.' : 0x'.dechex($dataHeaderRG['recType']).')');
                    }
                }
            } while (in_array($dataHeaderRG['recType'], $array));
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a persist object directory. Each persist object identifier specified MUST be unique in that persist object directory.
     * @link http://msdn.microsoft.com/en-us/library/dd952680(v=office.12).aspx
     * @param string $stream
     * @param integer $pos
     * @throws \Exception
     */
    private function readRecordPersistDirectoryAtom($stream, $pos)
    {
        $rHeader = $this->loadRecordHeader($stream, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0x0 || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_PERSISTDIRECTORYATOM) {
            throw new \Exception('File PowerPoint 97 in error (Location : PersistDirectoryAtom > RecordHeader).');
        }
        // rgPersistDirEntry
        // @link : http://msdn.microsoft.com/en-us/library/dd947347(v=office.12).aspx
        do {
            $data = self::getInt4d($stream, $pos);
            $pos += 4;
            $rHeader['recLen'] -= 4;
            //$persistId  = ($data >> 0) & bindec('11111111111111111111');
            $cPersist  = ($data >> 20) & bindec('111111111111');

            $rgPersistOffset = array();
            for ($inc = 0; $inc < $cPersist; $inc++) {
                $rgPersistOffset[] = self::getInt4d($stream, $pos);
                $pos += 4;
                $rHeader['recLen'] -= 4;
            }
        } while ($rHeader['recLen'] > 0);
        $this->rgPersistDirEntry = $rgPersistOffset;
    }

    /**
     * A container record that specifies information about the headers (1) and footers within a slide.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd904856(v=office.12).aspx
     */
    private function readRecordPerSlideHeadersFootersContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_HEADERSFOOTERS) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies whether a shape is a placeholder shape.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd923930(v=office.12).aspx
     */
    private function readRecordPlaceholderAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_PLACEHOLDERATOM && $data['recLen'] == 0x00000008) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a collection of re-color mappings for a metafile ([MS-WMF]).
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd904899(v=office.12).aspx
     */
    private function readRecordRecolorInfoAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_RECOLORINFOATOM) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies that a shape is a header or footerplaceholder shape.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd910800(v=office.12).aspx
     */
    private function readRecordRoundTripHFPlaceholder12Atom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_ROUNDTRIPHFPLACEHOLDER12ATOM && $data['recLen'] == 0x00000001) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a shape identifier.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd772926(v=office.12).aspx
     */
    private function readRecordRoundTripShapeId12Atom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_ROUNDTRIPSHAPEID12ATOM && $data['recLen'] == 0x00000004) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies information about a slide that synchronizes to a slide in a slide library.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd923801(v=office.12).aspx
     */
    private function readRecordRoundTripSlideSyncInfo12Container($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_ROUNDTRIPSLIDESYNCINFO12) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies shape-level Boolean flags.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd908949(v=office.12).aspx
     */
    private function readRecordShapeFlags10Atom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_SHAPEFLAGS10ATOM && $data['recLen'] == 0x00000001) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies shape-level Boolean flags.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd925824(v=office.12).aspx
     */
    private function readRecordShapeFlagsAtom($stream, $pos)
    {
        $arrayReturn = array(
                'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_SHAPEATOM && $data['recLen'] == 0x00000001) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies programmable tags with additional binary shape data.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd911033(v=office.12).aspx
     */
    private function readRecordShapeProgBinaryTagContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_PROGBINARYTAG) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies programmable tags with additional shape data.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd911266(v=office.12).aspx
     */
    private function readRecordShapeProgTagsContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_PROGTAGS) {
            // Record Header
            $arrayReturn['length'] += 8;

            $length = 0;
            do {
                $dataHeaderRG = $this->loadRecordHeader($stream, $pos + $arrayReturn['length'] + $length);
                switch ($dataHeaderRG['recType']) {
                    case self::RT_PROGBINARYTAG:
                        $dataRG = $this->readRecordShapeProgBinaryTagContainer($stream, $pos + $arrayReturn['length'] + $length);
                        $length += $dataRG['length'];
                        break;
                    //case self::RT_PROGSTRINGTAG:
                    default:
                        throw new \Exception('Feature not implemented (l.'.__LINE__.')');
                }
            } while ($length < $data['recLen']);
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies information about a slide.
     * @param string $stream
     * @param integer $pos
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd923801(v=office.12).aspx
     */
    private function readRecordSlideAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x2 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_SLIDEATOM) {
            // Record Header
            $arrayReturn['length'] += 8;
            // slideAtom > geom
            $arrayReturn['length'] += 4;
            // slideAtom > rgPlaceholderTypes
            $rgPlaceholderTypes = array();
            for ($inc = 0; $inc < 8; $inc++) {
                $rgPlaceholderTypes[] = self::getInt1d($this->streamPowerpointDocument, $pos);
                $arrayReturn['length'] += 1;
            }

            // slideAtom > masterIdRef
            $arrayReturn['length'] += 4;
            // slideAtom > notesIdRef
            $arrayReturn['length'] += 4;
            // slideAtom > slideFlags
            $arrayReturn['length'] += 2;
            // slideAtom > unused;
            $arrayReturn['length'] += 2;
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies a presentation slide or title master slide.
     * @param string $stream
     * @param int $pos
     * @link http://msdn.microsoft.com/en-us/library/dd946323(v=office.12).aspx
     */
    private function readRecordSlideContainer($stream, $pos)
    {
        // Core
        $this->oPhpPresentation->createSlide();
        $this->oPhpPresentation->setActiveSlideIndex($this->oPhpPresentation->getSlideCount() - 1);

        // *** slideAtom (32 bytes)
        $slideAtom = $this->readRecordSlideAtom($stream, $pos);
        if ($slideAtom['length'] == 0) {
            throw new \Exception('PowerPoint97 Reader : record SlideAtom');
        }
        $pos += $slideAtom['length'];

        // *** slideShowSlideInfoAtom (24 bytes)
        $slideShowInfoAtom = $this->readRecordSlideShowSlideInfoAtom($stream, $pos);
        $pos += $slideShowInfoAtom['length'];

        // *** perSlideHFContainer (variable) : optional
        $perSlideHFContainer = $this->readRecordPerSlideHeadersFootersContainer($stream, $pos);
        $pos += $perSlideHFContainer['length'];

        // *** rtSlideSyncInfo12 (variable) : optional
        $rtSlideSyncInfo12 = $this->readRecordRoundTripSlideSyncInfo12Container($stream, $pos);
        $pos += $rtSlideSyncInfo12['length'];

        // *** drawing (variable)
        $drawing = $this->readRecordDrawingContainer($stream, $pos);
        $pos += $drawing['length'];

        // *** slideSchemeColorSchemeAtom (40 bytes)
        $slideSchemeColorAtom = $this->readRecordSlideSchemeColorSchemeAtom($stream, $pos);
        if ($slideSchemeColorAtom['length'] == 0) {
            throw new \Exception('PowerPoint97 Reader : record SlideSchemeColorSchemeAtom');
        }
        $pos += $slideSchemeColorAtom['length'];

        // *** slideNameAtom (variable)
        $slideNameAtom = $this->readRecordSlideNameAtom($stream, $pos);
        $pos += $slideNameAtom['length'];

        // *** slideProgTagsContainer (variable).
        $slideProgTags = $this->readRecordSlideProgTagsContainer($stream, $pos);
        $pos += $slideProgTags['length'];

        // *** rgRoundTripSlide (variable)
    }

    /**
     * An atom record that specifies the name of a slide.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd906297(v=office.12).aspx
     */
    private function readRecordSlideNameAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
            'slideName' => '',
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x003 && $data['recType'] == self::RT_CSTRING && $data['recLen'] % 2 == 0) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $strLen = ($data['recLen'] / 2);
            for ($inc = 0; $inc < $strLen; $inc++) {
                $char = self::getInt2d($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $arrayReturn['slideName'] .= Text::chr($char);
            }
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a slide number metacharacter.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd945703(v=office.12).aspx
     */
    private function readRecordSlideNumberMCAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_SLIDENUMBERMETACHARATOM && $data['recLen'] == 0x00000004) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies programmable tags with additional slide data.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd951946(v=office.12).aspx
     */
    private function readRecordSlideProgTagsContainer($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0xF && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_PROGTAGS) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies the color scheme used by a slide.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd949420(v=office.12).aspx
     */
    private function readRecordSlideSchemeColorSchemeAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x001 && $data['recType'] == self::RT_COLORSCHEMEATOM && $data['recLen'] == 0x00000020) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $rgSchemeColor = array();
            for ($inc = 0; $inc <= 7; $inc++) {
                $rgSchemeColor[] = array(
                    'red' => self::getInt1d($stream, $pos + $arrayReturn['length'] + $inc * 4),
                    'green' => self::getInt1d($stream, $pos + $arrayReturn['length'] + $inc * 4 + 1),
                    'blue' => self::getInt1d($stream, $pos + $arrayReturn['length'] + $inc * 4 + 2),
                );
            }
            $arrayReturn['length'] += (8 * 4);
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies what transition effect to perform during a slide show, and how to advance to the next presentation slide.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd943408(v=office.12).aspx
     */
    private function readRecordSlideShowSlideInfoAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] == 0x0 && $data['recInstance'] == 0x000 && $data['recType'] == self::RT_SLIDESHOWSLIDEINFOATOM && $data['recLen'] == 0x00000010) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length;
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * UserEditAtom
     * @link http://msdn.microsoft.com/en-us/library/dd945746(v=office.12).aspx
     * @param string $stream
     * @param integer $pos
     * @throws \Exception
     */
    private function readRecordUserEditAtom($stream, $pos)
    {
        $rHeader = $this->loadRecordHeader($stream, $pos);
        $pos += 8;
        if ($rHeader['recVer'] != 0x0 || $rHeader['recInstance'] != 0x000 || $rHeader['recType'] != self::RT_USEREDITATOM || ($rHeader['recLen'] != 0x0000001C && $rHeader['recLen'] != 0x00000020)) {
            throw new \Exception('File PowerPoint 97 in error (Location : UserEditAtom > RecordHeader).');
        }

        // lastSlideIdRef
        $pos += 4;
        // version
        $pos += 2;

        // minorVersion
        $minorVersion = self::getInt1d($stream, $pos);
        $pos += 1;
        if ($minorVersion != 0x00) {
            throw new \Exception('File PowerPoint 97 in error (Location : UserEditAtom > minorVersion).');
        }

        // majorVersion
        $majorVersion = self::getInt1d($stream, $pos);
        $pos += 1;
        if ($majorVersion != 0x03) {
            throw new \Exception('File PowerPoint 97 in error (Location : UserEditAtom > majorVersion).');
        }

        // offsetLastEdit
        $pos += 4;
        // offsetPersistDirectory
        $this->offsetPersistDirectory = self::getInt4d($stream, $pos);
        $pos += 4;

        // docPersistIdRef
        $docPersistIdRef  = self::getInt4d($stream, $pos);
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
     * A structure that specifies the character-level formatting of a run of text.
     * @param string $stream
     * @param int $pos
     * @param int $strLenRT
     * @link https://msdn.microsoft.com/en-us/library/dd945870(v=office.12).aspx
     */
    private function readStructureTextCFRun($stream, $pos, $strLenRT)
    {
        $arrayReturn = array(
            'length' => 0,
            'strLenRT' => $strLenRT,
        );

        // rgTextCFRun
        $countRgTextCFRun = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['strLenRT'] -= $countRgTextCFRun;
        $arrayReturn['length'] += 4;
        $arrayReturn['partLength'] = $countRgTextCFRun;

        $masks = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;

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
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;

            $fontStyleFlags = array();
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

            $arrayReturn['bold'] = ($fontStyleFlags['bold'] == 1) ? true : false;
            $arrayReturn['italic'] = ($fontStyleFlags['italic'] == 1) ? true : false;
            $arrayReturn['underline'] = ($fontStyleFlags['underline'] == 1) ? true : false;
        }
        if ($masksData['typeface'] == 1) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayReturn['fontName'] = isset($this->arrayFonts[$data]) ? $this->arrayFonts[$data] : '';
        }
        if ($masksData['oldEATypeface'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['ansiTypeface'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['symbolTypeface'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['size'] == 1) {
            $arrayReturn['fontSize'] = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['color'] == 1) {
            $red = self::getInt1d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 1;
            $green = self::getInt1d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 1;
            $blue = self::getInt1d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 1;
            $index = self::getInt1d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 1;

            if ($index == 0xFE) {
                $strColor = str_pad(dechex($red), 2, STR_PAD_LEFT, '0');
                $strColor .= str_pad(dechex($green), 2, STR_PAD_LEFT, '0');
                $strColor .= str_pad(dechex($blue), 2, STR_PAD_LEFT, '0');

                $arrayReturn['color'] = new Color('FF'.$strColor);
            }
        }
        if ($masksData['position'] == 1) {
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }

        return $arrayReturn;
    }

    /**
     * A structure that specifies the paragraph-level formatting of a run of text.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd923535(v=office.12).aspx
     */
    private function readStructureTextPFRun($stream, $pos, $strLenRT)
    {
        $arrayReturn = array(
            'length' => 0,
            'strLenRT' => $strLenRT,
        );

        // rgTextPFRun
        $countRgTextPFRun = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['strLenRT'] -= $countRgTextPFRun;
        $arrayReturn['length'] += 4;

        // indent
        $arrayReturn['length'] += 2;

        $masks = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;

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
        if ($masksData['hasBullet'] == 1 || $masksData['bulletHasFont'] == 1  || $masksData['bulletHasColor'] == 1  || $masksData['bulletHasSize'] == 1) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;

            $bulletFlags['fHasBullet'] = ($data >> 0) & bindec('1');
            $bulletFlags['fBulletHasFont'] = ($data >> 1) & bindec('1');
            $bulletFlags['fBulletHasColor'] = ($data >> 2) & bindec('1');
            $bulletFlags['fBulletHasSize'] = ($data >> 3) & bindec('1');
        }
        if ($masksData['bulletChar'] == 1) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayReturn['bulletChar'] = chr($data);
        }
        if ($masksData['bulletFont'] == 1) {
            // $data = self::getInt2d($stream, $pos);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['bulletSize'] == 1) {
            // $data = self::getInt2d($stream, $pos);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['bulletColor'] == 1) {
            $red = self::getInt1d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 1;
            $green = self::getInt1d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 1;
            $blue = self::getInt1d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 1;
            $index = self::getInt1d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 1;

            if ($index == 0xFE) {
                $strColor = str_pad(dechex($red), 2, STR_PAD_LEFT, '0');
                $strColor .= str_pad(dechex($green), 2, STR_PAD_LEFT, '0');
                $strColor .= str_pad(dechex($blue), 2, STR_PAD_LEFT, '0');
            }
        }
        if ($masksData['align'] == 1) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            switch ($data) {
                case 0x0000:
                    $arrayReturn['alignH'] = Alignment::HORIZONTAL_LEFT;
                    break;
                case 0x0001:
                    $arrayReturn['alignH'] = Alignment::HORIZONTAL_CENTER;
                    break;
                case 0x0002:
                    $arrayReturn['alignH'] = Alignment::HORIZONTAL_RIGHT;
                    break;
                case 0x0003:
                    $arrayReturn['alignH'] = Alignment::HORIZONTAL_JUSTIFY;
                    break;
                case 0x0004:
                    $arrayReturn['alignH'] = Alignment::HORIZONTAL_DISTRIBUTED;
                    break;
                case 0x0005:
                    $arrayReturn['alignH'] = Alignment::HORIZONTAL_DISTRIBUTED;
                    break;
                case 0x0006:
                    $arrayReturn['alignH'] = Alignment::HORIZONTAL_JUSTIFY;
                    break;
                default:
                    break;
            }
        }
        if ($masksData['lineSpacing'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['spaceBefore'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['spaceAfter'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['leftMargin'] == 1) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayReturn['leftMargin'] = (int)round($data/6);
        }
        if ($masksData['indent'] == 1) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayReturn['indent'] = (int)round($data/6);
        }
        if ($masksData['defaultTabSize'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['tabStops'] == 1) {
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }
        if ($masksData['fontAlign'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['charWrap'] == 1 || $masksData['wordWrap'] == 1 || $masksData['overflow'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['textDirection'] == 1) {
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }

        return $arrayReturn;
    }

    /**
     * A structure that specifies language and spelling information for a run of text.
     * @param string $stream
     * @param integer $pos
     * @param string $strLenRT
     * @return array
     * @link https://msdn.microsoft.com/en-us/library/dd909603(v=office.12).aspx
     */
    private function readStructureTextSIRun($stream, $pos, $strLenRT)
    {
        $arrayReturn = array(
            'length' => 0,
            'strLenRT' => $strLenRT,
        );

        $arrayReturn['strLenRT'] -= self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;

        $data = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;
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
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $masksSpell = array();
            $masksSpell['error'] = ($data >> 0) & bindec('1');
            $masksSpell['clean'] = ($data >> 1) & bindec('1');
            $masksSpell['grammar'] = ($data >> 2) & bindec('1');
        }
        if ($masksData['lang'] == 1) {
            // $data = self::getInt2d($stream, $pos);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['altLang'] == 1) {
            // $data = self::getInt2d($stream, $pos);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fBidi'] == 1) {
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }
        if ($masksData['fPp10ext'] == 1) {
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }
        if ($masksData['smartTag'] == 1) {
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }

        return $arrayReturn;
    }

    /**
     * A structure that specifies tabbing, margins, and indentation for text.
     * @param string $stream
     * @param integer $pos
     * @link https://msdn.microsoft.com/en-us/library/dd922749(v=office.12).aspx
     */
    private function readStructureTextRuler($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;

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
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }
        if ($masksData['fDefaultTabSize'] == 1) {
            throw new \Exception('Feature not implemented (l.'.__LINE__.')');
        }
        if ($masksData['fTabStops'] == 1) {
            $count = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayTabStops = array();
            for ($inc = 0; $inc < $count; $inc++) {
                $position = self::getInt2d($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $type = self::getInt2d($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $arrayTabStops[] = array(
                    'position' => $position,
                    'type' => $type,
                );
            }
        }
        if ($masksData['fLeftMargin1'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fIndent1'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fLeftMargin2'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fIndent2'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fLeftMargin3'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fIndent3'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fLeftMargin4'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fIndent4'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fLeftMargin5'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if ($masksData['fIndent5'] == 1) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }

        return $arrayReturn;
    }

    /**
     * @param $stream
     * @param int $pos
     * @throws \Exception
     */
    private function readRecordNotesContainer($stream, $pos)
    {
        // notesAtom
        $notesAtom = $this->readRecordNotesAtom($stream, $pos);
        $pos += $notesAtom['length'];

        // drawing
        $drawing = $this->readRecordDrawingContainer($stream, $pos);
        $pos += $drawing['length'];

        // slideSchemeColorSchemeAtom
        // slideNameAtom
        // slideProgTagsContainer
        // rgNotesRoundTripAtom
    }

    /**
     * @param $stream
     * @param int $pos
     * @return array
     * @throws \Exception
     */
    private function readRecordNotesAtom($stream, $pos)
    {
        $arrayReturn = array(
            'length' => 0,
        );

        $data = $this->loadRecordHeader($stream, $pos);
        if ($data['recVer'] != 0x1 || $data['recInstance'] != 0x000 || $data['recType'] != self::RT_NOTESATOM || $data['recLen'] != 0x00000008) {
            throw new \Exception('File PowerPoint 97 in error (Location : NotesAtom > RecordHeader)');
        }
        // Record Header
        $arrayReturn['length'] += 8;
        // NotesAtom > slideIdRef
        $notesIdRef = self::getInt4d($stream, $pos + $arrayReturn['length']);
        if ($notesIdRef == -2147483648) {
            $notesIdRef = 0;
        }
        $this->currentNote = $notesIdRef;
        $arrayReturn['length'] += 4;

        // NotesAtom > slideFlags
        $arrayReturn['length'] += 2;
        // NotesAtom > unused
        $arrayReturn['length'] += 2;

        return $arrayReturn;
    }
}
