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

namespace PhpOffice\PhpPresentation\Reader;

use Exception;
use PhpOffice\Common\Microsoft\OLERead;
use PhpOffice\Common\Text;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\Exception\FeatureNotImplementedException;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Hyperlink;
use PhpOffice\PhpPresentation\Shape\Line;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Font;

/**
 * Serialized format reader.
 */
class PowerPoint97 implements ReaderInterface
{
    public const OFFICEARTBLIPEMF = 0xF01A;
    public const OFFICEARTBLIPWMF = 0xF01B;
    public const OFFICEARTBLIPPICT = 0xF01C;
    public const OFFICEARTBLIPJPG = 0xF01D;
    public const OFFICEARTBLIPPNG = 0xF01E;
    public const OFFICEARTBLIPDIB = 0xF01F;
    public const OFFICEARTBLIPTIFF = 0xF029;
    public const OFFICEARTBLIPJPEG = 0xF02A;

    /**
     * @see http://msdn.microsoft.com/en-us/library/dd945336(v=office.12).aspx
     */
    public const RT_ANIMATIONINFO = 0x1014;
    public const RT_ANIMATIONINFOATOM = 0x0FF1;
    public const RT_BINARYTAGDATABLOB = 0x138B;
    public const RT_BLIPCOLLECTION9 = 0x07F8;
    public const RT_BLIPENTITY9ATOM = 0x07F9;
    public const RT_BOOKMARKCOLLECTION = 0x07E3;
    public const RT_BOOKMARKENTITYATOM = 0x0FD0;
    public const RT_BOOKMARKSEEDATOM = 0x07E9;
    public const RT_BROADCASTDOCINFO9 = 0x177E;
    public const RT_BROADCASTDOCINFO9ATOM = 0x177F;
    public const RT_BUILDATOM = 0x2B03;
    public const RT_BUILDLIST = 0x2B02;
    public const RT_CHARTBUILD = 0x2B04;
    public const RT_CHARTBUILDATOM = 0x2B05;
    public const RT_COLORSCHEMEATOM = 0x07F0;
    public const RT_COMMENT10 = 0x2EE0;
    public const RT_COMMENT10ATOM = 0x2EE1;
    public const RT_COMMENTINDEX10 = 0x2EE4;
    public const RT_COMMENTINDEX10ATOM = 0x2EE5;
    public const RT_CRYPTSESSION10CONTAINER = 0x2F14;
    public const RT_CURRENTUSERATOM = 0x0FF6;
    public const RT_CSTRING = 0x0FBA;
    public const RT_DATETIMEMETACHARATOM = 0x0FF7;
    public const RT_DEFAULTRULERATOM = 0x0FAB;
    public const RT_DOCROUTINGSLIPATOM = 0x0406;
    public const RT_DIAGRAMBUILD = 0x2B06;
    public const RT_DIAGRAMBUILDATOM = 0x2B07;
    public const RT_DIFF10 = 0x2EED;
    public const RT_DIFF10ATOM = 0x2EEE;
    public const RT_DIFFTREE10 = 0x2EEC;
    public const RT_DOCTOOLBARSTATES10ATOM = 0x36B1;
    public const RT_DOCUMENT = 0x03E8;
    public const RT_DOCUMENTATOM = 0x03E9;
    public const RT_DRAWING = 0x040C;
    public const RT_DRAWINGGROUP = 0x040B;
    public const RT_ENDDOCUMENTATOM = 0x03EA;
    public const RT_EXTERNALAVIMOVIE = 0x1006;
    public const RT_EXTERNALCDAUDIO = 0x100E;
    public const RT_EXTERNALCDAUDIOATOM = 0x1012;
    public const RT_EXTERNALHYPERLINK = 0x0FD7;
    public const RT_EXTERNALHYPERLINK9 = 0x0FE4;
    public const RT_EXTERNALHYPERLINKATOM = 0x0FD3;
    public const RT_EXTERNALHYPERLINKFLAGSATOM = 0x1018;
    public const RT_EXTERNALMCIMOVIE = 0x1007;
    public const RT_EXTERNALMEDIAATOM = 0x1004;
    public const RT_EXTERNALMIDIAUDIO = 0x100D;
    public const RT_EXTERNALOBJECTLIST = 0x0409;
    public const RT_EXTERNALOBJECTLISTATOM = 0x040A;
    public const RT_EXTERNALOBJECTREFATOM = 0x0BC1;
    public const RT_EXTERNALOLECONTROL = 0x0FEE;
    public const RT_EXTERNALOLECONTROLATOM = 0x0FFB;
    public const RT_EXTERNALOLEEMBED = 0x0FCC;
    public const RT_EXTERNALOLEEMBEDATOM = 0x0FCD;
    public const RT_EXTERNALOLELINK = 0x0FCE;
    public const RT_EXTERNALOLELINKATOM = 0x0FD1;
    public const RT_EXTERNALOLEOBJECTATOM = 0x0FC3;
    public const RT_EXTERNALOLEOBJECTSTG = 0x1011;
    public const RT_EXTERNALVIDEO = 0x1005;
    public const RT_EXTERNALWAVAUDIOEMBEDDED = 0x100F;
    public const RT_EXTERNALWAVAUDIOEMBEDDEDATOM = 0x1013;
    public const RT_EXTERNALWAVAUDIOLINK = 0x1010;
    public const RT_ENVELOPEDATA9ATOM = 0x1785;
    public const RT_ENVELOPEFLAGS9ATOM = 0x1784;
    public const RT_ENVIRONMENT = 0x03F2;
    public const RT_FONTCOLLECTION = 0x07D5;
    public const RT_FONTCOLLECTION10 = 0x07D6;
    public const RT_FONTEMBEDDATABLOB = 0x0FB8;
    public const RT_FONTEMBEDFLAGS10ATOM = 0x32C8;
    public const RT_FILTERPRIVACYFLAGS10ATOM = 0x36B0;
    public const RT_FONTENTITYATOM = 0x0FB7;
    public const RT_FOOTERMETACHARATOM = 0x0FFA;
    public const RT_GENERICDATEMETACHARATOM = 0x0FF8;
    public const RT_GRIDSPACING10ATOM = 0x040D;
    public const RT_GUIDEATOM = 0x03FB;
    public const RT_HANDOUT = 0x0FC9;
    public const RT_HASHCODEATOM = 0x2B00;
    public const RT_HEADERSFOOTERS = 0x0FD9;
    public const RT_HEADERSFOOTERSATOM = 0x0FDA;
    public const RT_HEADERMETACHARATOM = 0x0FF9;
    public const RT_HTMLDOCINFO9ATOM = 0x177B;
    public const RT_HTMLPUBLISHINFOATOM = 0x177C;
    public const RT_HTMLPUBLISHINFO9 = 0x177D;
    public const RT_INTERACTIVEINFO = 0x0FF2;
    public const RT_INTERACTIVEINFOATOM = 0x0FF3;
    public const RT_KINSOKU = 0x0FC8;
    public const RT_KINSOKUATOM = 0x0FD2;
    public const RT_LEVELINFOATOM = 0x2B0A;
    public const RT_LINKEDSHAPE10ATOM = 0x2EE6;
    public const RT_LINKEDSLIDE10ATOM = 0x2EE7;
    public const RT_LIST = 0x07D0;
    public const RT_MAINMASTER = 0x03F8;
    public const RT_MASTERTEXTPROPATOM = 0x0FA2;
    public const RT_METAFILE = 0x0FC1;
    public const RT_NAMEDSHOW = 0x0411;
    public const RT_NAMEDSHOWS = 0x0410;
    public const RT_NAMEDSHOWSLIDESATOM = 0x0412;
    public const RT_NORMALVIEWSETINFO9 = 0x0414;
    public const RT_NORMALVIEWSETINFO9ATOM = 0x0415;
    public const RT_NOTES = 0x03F0;
    public const RT_NOTESATOM = 0x03F1;
    public const RT_NOTESTEXTVIEWINFO9 = 0x0413;
    public const RT_OUTLINETEXTPROPS9 = 0x0FAE;
    public const RT_OUTLINETEXTPROPS10 = 0x0FB3;
    public const RT_OUTLINETEXTPROPS11 = 0x0FB5;
    public const RT_OUTLINETEXTPROPSHEADER9ATOM = 0x0FAF;
    public const RT_OUTLINETEXTREFATOM = 0x0F9E;
    public const RT_OUTLINEVIEWINFO = 0x0407;
    public const RT_PERSISTDIRECTORYATOM = 0x1772;
    public const RT_PARABUILD = 0x2B08;
    public const RT_PARABUILDATOM = 0x2B09;
    public const RT_PHOTOALBUMINFO10ATOM = 0x36B2;
    public const RT_PLACEHOLDERATOM = 0x0BC3;
    public const RT_PRESENTATIONADVISORFLAGS9ATOM = 0x177A;
    public const RT_PRINTOPTIONSATOM = 0x1770;
    public const RT_PROGBINARYTAG = 0x138A;
    public const RT_PROGSTRINGTAG = 0x1389;
    public const RT_PROGTAGS = 0x1388;
    public const RT_RECOLORINFOATOM = 0x0FE7;
    public const RT_RTFDATETIMEMETACHARATOM = 0x1015;
    public const RT_ROUNDTRIPANIMATIONATOM12ATOM = 0x2B0B;
    public const RT_ROUNDTRIPANIMATIONHASHATOM12ATOM = 0x2B0D;
    public const RT_ROUNDTRIPCOLORMAPPING12ATOM = 0x040F;
    public const RT_ROUNDTRIPCOMPOSITEMASTERID12ATOM = 0x041D;
    public const RT_ROUNDTRIPCONTENTMASTERID12ATOM = 0x0422;
    public const RT_ROUNDTRIPCONTENTMASTERINFO12ATOM = 0x041E;
    public const RT_ROUNDTRIPCUSTOMTABLESTYLES12ATOM = 0x0428;
    public const RT_ROUNDTRIPDOCFLAGS12ATOM = 0x0425;
    public const RT_ROUNDTRIPHEADERFOOTERDEFAULTS12ATOM = 0x0424;
    public const RT_ROUNDTRIPHFPLACEHOLDER12ATOM = 0x0420;
    public const RT_ROUNDTRIPNEWPLACEHOLDERID12ATOM = 0x0BDD;
    public const RT_ROUNDTRIPNOTESMASTERTEXTSTYLES12ATOM = 0x0427;
    public const RT_ROUNDTRIPOARTTEXTSTYLES12ATOM = 0x0423;
    public const RT_ROUNDTRIPORIGINALMAINMASTERID12ATOM = 0x041C;
    public const RT_ROUNDTRIPSHAPECHECKSUMFORCL12ATOM = 0x0426;
    public const RT_ROUNDTRIPSHAPEID12ATOM = 0x041F;
    public const RT_ROUNDTRIPSLIDESYNCINFO12 = 0x3714;
    public const RT_ROUNDTRIPSLIDESYNCINFOATOM12 = 0x3715;
    public const RT_ROUNDTRIPTHEME12ATOM = 0x040E;
    public const RT_SHAPEATOM = 0x0BDB;
    public const RT_SHAPEFLAGS10ATOM = 0x0BDC;
    public const RT_SLIDE = 0x03EE;
    public const RT_SLIDEATOM = 0x03EF;
    public const RT_SLIDEFLAGS10ATOM = 0x2EEA;
    public const RT_SLIDELISTENTRY10ATOM = 0x2EF0;
    public const RT_SLIDELISTTABLE10 = 0x2EF1;
    public const RT_SLIDELISTWITHTEXT = 0x0FF0;
    public const RT_SLIDELISTTABLESIZE10ATOM = 0x2EEF;
    public const RT_SLIDENUMBERMETACHARATOM = 0x0FD8;
    public const RT_SLIDEPERSISTATOM = 0x03F3;
    public const RT_SLIDESHOWDOCINFOATOM = 0x0401;
    public const RT_SLIDESHOWSLIDEINFOATOM = 0x03F9;
    public const RT_SLIDETIME10ATOM = 0x2EEB;
    public const RT_SLIDEVIEWINFO = 0x03FA;
    public const RT_SLIDEVIEWINFOATOM = 0x03FE;
    public const RT_SMARTTAGSTORE11CONTAINER = 0x36B3;
    public const RT_SOUND = 0x07E6;
    public const RT_SOUNDCOLLECTION = 0x07E4;
    public const RT_SOUNDCOLLECTIONATOM = 0x07E5;
    public const RT_SOUNDDATABLOB = 0x07E7;
    public const RT_SORTERVIEWINFO = 0x0408;
    public const RT_STYLETEXTPROPATOM = 0x0FA1;
    public const RT_STYLETEXTPROP10ATOM = 0x0FB1;
    public const RT_STYLETEXTPROP11ATOM = 0x0FB6;
    public const RT_STYLETEXTPROP9ATOM = 0x0FAC;
    public const RT_SUMMARY = 0x0402;
    public const RT_TEXTBOOKMARKATOM = 0x0FA7;
    public const RT_TEXTBYTESATOM = 0x0FA8;
    public const RT_TEXTCHARFORMATEXCEPTIONATOM = 0x0FA4;
    public const RT_TEXTCHARSATOM = 0x0FA0;
    public const RT_TEXTDEFAULTS10ATOM = 0x0FB4;
    public const RT_TEXTDEFAULTS9ATOM = 0x0FB0;
    public const RT_TEXTHEADERATOM = 0x0F9F;
    public const RT_TEXTINTERACTIVEINFOATOM = 0x0FDF;
    public const RT_TEXTMASTERSTYLEATOM = 0x0FA3;
    public const RT_TEXTMASTERSTYLE10ATOM = 0x0FB2;
    public const RT_TEXTMASTERSTYLE9ATOM = 0x0FAD;
    public const RT_TEXTPARAGRAPHFORMATEXCEPTIONATOM = 0x0FA5;
    public const RT_TEXTRULERATOM = 0x0FA6;
    public const RT_TEXTSPECIALINFOATOM = 0x0FAA;
    public const RT_TEXTSPECIALINFODEFAULTATOM = 0x0FA9;
    public const RT_TIMEANIMATEBEHAVIOR = 0xF134;
    public const RT_TIMEANIMATEBEHAVIORCONTAINER = 0xF12B;
    public const RT_TIMEANIMATIONVALUE = 0xF143;
    public const RT_TIMEANIMATIONVALUELIST = 0xF13F;
    public const RT_TIMEBEHAVIOR = 0xF133;
    public const RT_TIMEBEHAVIORCONTAINER = 0xF12A;
    public const RT_TIMECOLORBEHAVIOR = 0xF135;
    public const RT_TIMECOLORBEHAVIORCONTAINER = 0xF12C;
    public const RT_TIMECLIENTVISUALELEMENT = 0xF13C;
    public const RT_TIMECOMMANDBEHAVIOR = 0xF13B;
    public const RT_TIMECOMMANDBEHAVIORCONTAINER = 0xF132;
    public const RT_TIMECONDITION = 0xF128;
    public const RT_TIMECONDITIONCONTAINER = 0xF125;
    public const RT_TIMEEFFECTBEHAVIOR = 0xF136;
    public const RT_TIMEEFFECTBEHAVIORCONTAINER = 0xF12D;
    public const RT_TIMEEXTTIMENODECONTAINER = 0xF144;
    public const RT_TIMEITERATEDATA = 0xF140;
    public const RT_TIMEMODIFIER = 0xF129;
    public const RT_TIMEMOTIONBEHAVIOR = 0xF137;
    public const RT_TIMEMOTIONBEHAVIORCONTAINER = 0xF12E;
    public const RT_TIMENODE = 0xF127;
    public const RT_TIMEPROPERTYLIST = 0xF13D;
    public const RT_TIMEROTATIONBEHAVIOR = 0xF138;
    public const RT_TIMEROTATIONBEHAVIORCONTAINER = 0xF12F;
    public const RT_TIMESCALEBEHAVIOR = 0xF139;
    public const RT_TIMESCALEBEHAVIORCONTAINER = 0xF130;
    public const RT_TIMESEQUENCEDATA = 0xF141;
    public const RT_TIMESETBEHAVIOR = 0xF13A;
    public const RT_TIMESETBEHAVIORCONTAINER = 0xF131;
    public const RT_TIMESUBEFFECTCONTAINER = 0xF145;
    public const RT_TIMEVARIANT = 0xF142;
    public const RT_TIMEVARIANTLIST = 0xF13E;
    public const RT_USEREDITATOM = 0x0FF5;
    public const RT_VBAINFO = 0x03FF;
    public const RT_VBAINFOATOM = 0x0400;
    public const RT_VIEWINFOATOM = 0x03FD;
    public const RT_VISUALPAGEATOM = 0x2B01;
    public const RT_VISUALSHAPEATOM = 0x2AFB;

    /**
     * @see http://msdn.microsoft.com/en-us/library/dd926394(v=office.12).aspx
     */
    public const SL_BIGOBJECT = 0x0000000F;
    public const SL_BLANK = 0x00000010;
    public const SL_COLUMNTWOROWS = 0x0000000A;
    public const SL_FOUROBJECTS = 0x0000000E;
    public const SL_MASTERTITLE = 0x00000002;
    public const SL_TITLEBODY = 0x00000001;
    public const SL_TITLEONLY = 0x00000007;
    public const SL_TITLESLIDE = 0x00000000;
    public const SL_TWOCOLUMNS = 0x00000008;
    public const SL_TWOCOLUMNSROW = 0x0000000D;
    public const SL_TWOROWS = 0x00000009;
    public const SL_TWOROWSCOLUMN = 0x0000000B;
    public const SL_VERTICALTITLEBODY = 0x00000011;
    public const SL_VERTICALTWOROWS = 0x00000012;

    /**
     * Array with Fonts.
     *
     * @var array<int, string>
     */
    private $arrayFonts = [];

    /**
     * Array with Hyperlinks.
     *
     * @var array<int, array<string, string>>
     */
    private $arrayHyperlinks = [];

    /**
     * Array with Notes.
     *
     * @var array<int, int>
     */
    private $arrayNotes = [];

    /**
     * Array with Pictures.
     *
     * @var array<int, string>
     */
    private $arrayPictures = [];

    /**
     * Offset (in bytes) from the beginning of the PowerPoint Document Stream to the UserEditAtom record for the most recent user edit.
     *
     * @var int
     */
    private $offsetToCurrentEdit;

    /**
     * A structure that specifies a compressed table of sequential persist object identifiers and stream offsets to associated persist objects.
     *
     * @var array<int, int>
     */
    private $rgPersistDirEntry;

    /**
     * Offset (in bytes) from the beginning of the PowerPoint Document Stream to the PersistDirectoryAtom record for this user edit.
     *
     * @var int
     */
    private $offsetPersistDirectory;

    /**
     * Output Object.
     *
     * @var PhpPresentation
     */
    private $oPhpPresentation;

    /**
     * @var null|Group
     */
    private $oCurrentGroup;

    /**
     * @var bool
     */
    private $bFirstShapeGroup = false;

    /**
     * Stream "Powerpoint Document".
     *
     * @var string
     */
    private $streamPowerpointDocument;

    /**
     * Stream "Current User".
     *
     * @var string
     */
    private $streamCurrentUser;

    /**
     * Stream "Pictures".
     *
     * @var string
     */
    private $streamPictures;

    /**
     * @var int
     */
    private $inMainType;

    /**
     * @var null|int
     */
    private $currentNote;

    /**
     * @var null|string
     */
    private $filename;

    /**
     * @var bool
     */
    protected $loadImages = true;

    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     */
    public function canRead(string $pFilename): bool
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     */
    public function fileSupportsUnserializePhpPresentation(string $pFilename = ''): bool
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new FileNotFoundException($pFilename);
        }

        try {
            // Use ParseXL for the hard work.
            $ole = new OLERead();
            // get excel data
            $ole->read($pFilename);

            return true;
        } catch (Exception $e) {
            return false;
        }
    }

    /**
     * Loads PhpPresentation Serialized file.
     */
    public function load(string $pFilename, int $flags = 0): PhpPresentation
    {
        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new InvalidFileFormatException($pFilename, self::class);
        }

        $this->filename = $pFilename;

        $this->loadImages = !((bool) ($flags & self::SKIP_IMAGES));

        return $this->loadFile();
    }

    /**
     * Load PhpPresentation Serialized file.
     */
    private function loadFile(): PhpPresentation
    {
        $this->oPhpPresentation = new PhpPresentation();
        $this->oPhpPresentation->removeSlideByIndex();

        // Read OLE Blocks
        $this->loadOLE();
        // Read pictures in the Pictures Stream
        $this->loadPicturesStream();
        // Read information in the Current User Stream
        $this->loadCurrentUserStream();
        // Read information in the PowerPoint Document Stream
        $this->loadPowerpointDocumentStream();

        return $this->oPhpPresentation;
    }

    /**
     * Read OLE Part.
     */
    private function loadOLE(): void
    {
        // OLE reader
        $oOLE = new OLERead();
        $oOLE->read($this->filename);

        // PowerPoint Document Stream
        $this->streamPowerpointDocument = $oOLE->getStream($oOLE->powerpointDocument);

        // Current User Stream
        $this->streamCurrentUser = $oOLE->getStream($oOLE->currentUser);

        // Get pictures data
        $this->streamPictures = $oOLE->getStream($oOLE->pictures);
    }

    /**
     * Stream Pictures.
     *
     * @see http://msdn.microsoft.com/en-us/library/dd920746(v=office.12).aspx
     */
    private function loadPicturesStream(): void
    {
        $stream = $this->streamPictures;

        $pos = 0;
        do {
            $arrayRH = $this->loadRecordHeader($stream, $pos);
            $pos += 8;
            $readSuccess = false;
            if (0x00 == $arrayRH['recVer'] && (0xF007 == $arrayRH['recType'] || ($arrayRH['recType'] >= 0xF018 && $arrayRH['recType'] <= 0xF117))) {
                //@link : http://msdn.microsoft.com/en-us/library/dd950560(v=office.12).aspx
                if (0xF007 == $arrayRH['recType']) {
                    // OfficeArtFBSE
                    throw new FeatureNotImplementedException();
                }
                // $arrayRH['recType'] >= 0xF018 && $arrayRH['recType'] <= 0xF117
                $arrayRecord = $this->readRecordOfficeArtBlip($stream, $pos - 8);
                if ($arrayRecord['length'] > 0) {
                    $pos += $arrayRecord['length'];
                    $this->arrayPictures[] = $arrayRecord['picture'];
                }
                $readSuccess = true;
            }
        } while (true === $readSuccess);
    }

    /**
     * Stream Current User.
     *
     * @see http://msdn.microsoft.com/en-us/library/dd908567(v=office.12).aspx
     */
    private function loadCurrentUserStream(): void
    {
        $pos = 0;

        /**
         * CurrentUserAtom : http://msdn.microsoft.com/en-us/library/dd948895(v=office.12).aspx.
         */
        // RecordHeader : http://msdn.microsoft.com/en-us/library/dd926377(v=office.12).aspx
        $rHeader = $this->loadRecordHeader($this->streamCurrentUser, $pos);
        $pos += 8;
        if (0x0 != $rHeader['recVer'] || 0x000 != $rHeader['recInstance'] || self::RT_CURRENTUSERATOM != $rHeader['recType']) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : CurrentUserAtom > RecordHeader');
        }

        // Size
        $size = self::getInt4d($this->streamCurrentUser, $pos);
        $pos += 4;
        if (0x00000014 != $size) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : CurrentUserAtom > Size');
        }

        // headerToken
        $headerToken = self::getInt4d($this->streamCurrentUser, $pos);
        $pos += 4;
        if (0xF3D1C4DF == $headerToken) {
            // Encrypted file
            throw new FeatureNotImplementedException();
        }

        // offsetToCurrentEdit
        $this->offsetToCurrentEdit = self::getInt4d($this->streamCurrentUser, $pos);
        $pos += 4;

        // lenUserName
        $lenUserName = self::getInt2d($this->streamCurrentUser, $pos);
        $pos += 2;
        if ($lenUserName > 255) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : CurrentUserAtom > lenUserName');
        }

        // docFileVersion
        $docFileVersion = self::getInt2d($this->streamCurrentUser, $pos);
        $pos += 2;
        if (0x03F4 != $docFileVersion) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : CurrentUserAtom > docFileVersion');
        }

        // majorVersion
        $majorVersion = self::getInt1d($this->streamCurrentUser, $pos);
        ++$pos;
        if (0x03 != $majorVersion) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : CurrentUserAtom > majorVersion');
        }

        // minorVersion
        $minorVersion = self::getInt1d($this->streamCurrentUser, $pos);
        ++$pos;
        if (0x00 != $minorVersion) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : CurrentUserAtom > minorVersion');
        }

        // unused
        $pos += 2;

        // ansiUserName
        // $ansiUserName = '';
        do {
            $char = self::getInt1d($this->streamCurrentUser, $pos);
            if (($char >= 0x00 && $char <= 0x1F) || ($char >= 0x7F && $char <= 0x9F)) {
                $char = false;
            } else {
                // $ansiUserName .= chr($char);
                ++$pos;
            }
        } while (false !== $char);

        // relVersion
        $relVersion = self::getInt4d($this->streamCurrentUser, $pos);
        $pos += 4;
        if (0x00000008 != $relVersion && 0x00000009 != $relVersion) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : CurrentUserAtom > relVersion');
        }

        // unicodeUserName
        // $unicodeUserName = '';
        for ($inc = 0; $inc < $lenUserName; ++$inc) {
            $char = self::getInt2d($this->streamCurrentUser, $pos);
            if (($char >= 0x00 && $char <= 0x1F) || ($char >= 0x7F && $char <= 0x9F)) {
                break;
            }
            // $unicodeUserName .= chr($char);
            $pos += 2;
        }
    }

    /**
     * Stream Powerpoint Document.
     *
     * @see http://msdn.microsoft.com/en-us/library/dd921564(v=office.12).aspx
     */
    private function loadPowerpointDocumentStream(): void
    {
        $this->readRecordUserEditAtom($this->streamPowerpointDocument, $this->offsetToCurrentEdit);

        $this->readRecordPersistDirectoryAtom($this->streamPowerpointDocument, $this->offsetPersistDirectory);

        foreach ($this->rgPersistDirEntry as $offsetDir) {
            $pos = $offsetDir;

            $rHeader = $this->loadRecordHeader($this->streamPowerpointDocument, $pos);
            $pos += 8;
            $this->inMainType = $rHeader['recType'];
            $this->currentNote = null;
            switch ($rHeader['recType']) {
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
                    break;
            }
        }
    }

    /**
     * Read a record header.
     *
     * @return array<string, int>
     */
    private function loadRecordHeader(string $stream, int $pos): array
    {
        $rec = self::getInt2d($stream, $pos);
        $recType = self::getInt2d($stream, $pos + 2);
        $recLen = self::getInt4d($stream, $pos + 4);

        return [
            'recVer' => ($rec >> 0) & bindec('1111'),
            'recInstance' => ($rec >> 4) & bindec('111111111111'),
            'recType' => $recType,
            'recLen' => $recLen,
        ];
    }

    /**
     * Read 8-bit unsigned integer.
     */
    public static function getInt1d(string $data, int $pos): int
    {
        return ord($data[$pos]);
    }

    /**
     * Read 16-bit unsigned integer.
     */
    public static function getInt2d(string $data, int $pos): int
    {
        return ord($data[$pos]) | (ord($data[$pos + 1]) << 8);
    }

    /**
     * Read 32-bit signed integer.
     */
    public static function getInt4d(string $data, int $pos): int
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

        return ord($data[$pos]) | (ord($data[$pos + 1]) << 8) | (ord($data[$pos + 2]) << 16) | $ord24;
    }

    /**
     * A container record that specifies the animation and sound information for a shape.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd772900(v=office.12).aspx
     */
    private function readRecordAnimationInfoContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_ANIMATIONINFO == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;

            // animationAtom
            // animationSound
            throw new FeatureNotImplementedException();
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies information about the document.
     *
     * @see http://msdn.microsoft.com/en-us/library/dd947357(v=office.12).aspx
     */
    private function readRecordDocumentContainer(string $stream, int $pos): void
    {
        $documentAtom = $this->loadRecordHeader($stream, $pos);
        $pos += 8;
        if (0x1 != $documentAtom['recVer'] || 0x000 != $documentAtom['recInstance'] || self::RT_DOCUMENTATOM != $documentAtom['recType']) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : RTDocument > DocumentAtom');
        }
        $pos += $documentAtom['recLen'];

        $exObjList = $this->loadRecordHeader($stream, $pos);
        if (0xF == $exObjList['recVer'] && 0x000 == $exObjList['recInstance'] && self::RT_EXTERNALOBJECTLIST == $exObjList['recType']) {
            $pos += 8;
            // exObjListAtom > rh
            $exObjListAtom = $this->loadRecordHeader($stream, $pos);
            if (0x0 != $exObjListAtom['recVer'] || 0x000 != $exObjListAtom['recInstance'] || self::RT_EXTERNALOBJECTLISTATOM != $exObjListAtom['recType'] || 0x00000004 != $exObjListAtom['recLen']) {
                throw new InvalidFileFormatException($this->filename, self::class, 'Location : RTDocument > DocumentAtom > exObjList > exObjListAtom');
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
                        if (0x0 != $exHyperlinkAtom['recVer'] || 0x000 != $exHyperlinkAtom['recInstance'] || self::RT_EXTERNALHYPERLINKATOM != $exHyperlinkAtom['recType'] || 0x00000004 != $exHyperlinkAtom['recLen']) {
                            throw new InvalidFileFormatException($this->filename, self::class, 'Location : RTDocument > DocumentAtom > exObjList > rgChildRec > RT_ExternalHyperlink');
                        }
                        $pos += 8;
                        $exObjList['recLen'] -= 8;
                        // exHyperlinkAtom > exHyperlinkId
                        $exHyperlinkId = self::getInt4d($stream, $pos);
                        $pos += 4;
                        $exObjList['recLen'] -= 4;

                        $this->arrayHyperlinks[$exHyperlinkId] = [];
                        // friendlyNameAtom
                        $friendlyNameAtom = $this->loadRecordHeader($stream, $pos);
                        if (0x0 == $friendlyNameAtom['recVer'] && 0x000 == $friendlyNameAtom['recInstance'] && self::RT_CSTRING == $friendlyNameAtom['recType'] && $friendlyNameAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $this->arrayHyperlinks[$exHyperlinkId]['text'] = '';
                            for ($inc = 0; $inc < ($friendlyNameAtom['recLen'] / 2); ++$inc) {
                                $char = self::getInt2d($stream, $pos);
                                $pos += 2;
                                $exObjList['recLen'] -= 2;
                                $this->arrayHyperlinks[$exHyperlinkId]['text'] .= chr($char);
                            }
                        }
                        // targetAtom
                        $targetAtom = $this->loadRecordHeader($stream, $pos);
                        if (0x0 == $targetAtom['recVer'] && 0x001 == $targetAtom['recInstance'] && self::RT_CSTRING == $targetAtom['recType'] && $targetAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $this->arrayHyperlinks[$exHyperlinkId]['url'] = '';
                            for ($inc = 0; $inc < ($targetAtom['recLen'] / 2); ++$inc) {
                                $char = self::getInt2d($stream, $pos);
                                $pos += 2;
                                $exObjList['recLen'] -= 2;
                                $this->arrayHyperlinks[$exHyperlinkId]['url'] .= chr($char);
                            }
                        }
                        // locationAtom
                        $locationAtom = $this->loadRecordHeader($stream, $pos);
                        if (0x0 == $locationAtom['recVer'] && 0x003 == $locationAtom['recInstance'] && self::RT_CSTRING == $locationAtom['recType'] && $locationAtom['recLen'] % 2 == 0) {
                            $pos += 8;
                            $exObjList['recLen'] -= 8;
                            $string = '';
                            for ($inc = 0; $inc < ($locationAtom['recLen'] / 2); ++$inc) {
                                $char = self::getInt2d($stream, $pos);
                                $pos += 2;
                                $exObjList['recLen'] -= 2;
                                $string .= chr($char);
                            }
                        }

                        break;
                    default:
                        // var_dump(dechex((int) $childRec['recType']));
                        throw new FeatureNotImplementedException();
                }
            } while ($exObjList['recLen'] > 0);
        }

        //@link : http://msdn.microsoft.com/en-us/library/dd907813(v=office.12).aspx
        $documentTextInfo = $this->loadRecordHeader($stream, $pos);
        if (0xF == $documentTextInfo['recVer'] && 0x000 == $documentTextInfo['recInstance'] && self::RT_ENVIRONMENT == $documentTextInfo['recType']) {
            $pos += 8;
            //@link : http://msdn.microsoft.com/en-us/library/dd952717(v=office.12).aspx
            $kinsoku = $this->loadRecordHeader($stream, $pos);
            if (0xF == $kinsoku['recVer'] && 0x002 == $kinsoku['recInstance'] && self::RT_KINSOKU == $kinsoku['recType']) {
                $pos += 8;
                $pos += $kinsoku['recLen'];
            }

            //@link : http://msdn.microsoft.com/en-us/library/dd948152(v=office.12).aspx
            $fontCollection = $this->loadRecordHeader($stream, $pos);
            if (0xF == $fontCollection['recVer'] && 0x000 == $fontCollection['recInstance'] && self::RT_FONTCOLLECTION == $fontCollection['recType']) {
                $pos += 8;
                do {
                    $fontEntityAtom = $this->loadRecordHeader($stream, $pos);
                    $pos += 8;
                    $fontCollection['recLen'] -= 8;
                    if (0x0 != $fontEntityAtom['recVer'] || $fontEntityAtom['recInstance'] > 128 || self::RT_FONTENTITYATOM != $fontEntityAtom['recType']) {
                        throw new InvalidFileFormatException($this->filename, self::class, 'Location : RTDocument > RT_Environment > RT_FontCollection > RT_FontEntityAtom');
                    }
                    $string = '';
                    for ($inc = 0; $inc < 32; ++$inc) {
                        $char = self::getInt2d($stream, $pos);
                        $pos += 2;
                        $fontCollection['recLen'] -= 2;
                        $string .= chr($char);
                    }
                    $this->arrayFonts[] = $string;

                    // lfCharSet (1 byte)
                    ++$pos;
                    --$fontCollection['recLen'];

                    // fEmbedSubsetted (1 bit)
                    // unused (7 bits)
                    ++$pos;
                    --$fontCollection['recLen'];

                    // rasterFontType (1 bit)
                    // deviceFontType (1 bit)
                    // truetypeFontType (1 bit)
                    // fNoFontSubstitution (1 bit)
                    // reserved (4 bits)
                    ++$pos;
                    --$fontCollection['recLen'];

                    // lfPitchAndFamily (1 byte)
                    ++$pos;
                    --$fontCollection['recLen'];

                    $fontEmbedData1 = $this->loadRecordHeader($stream, $pos);
                    if (0x0 == $fontEmbedData1['recVer'] && $fontEmbedData1['recInstance'] >= 0x000 && $fontEmbedData1['recInstance'] <= 0x003 && self::RT_FONTEMBEDDATABLOB == $fontEmbedData1['recType']) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData1['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData1['recLen'];
                    }

                    $fontEmbedData2 = $this->loadRecordHeader($stream, $pos);
                    if (0x0 == $fontEmbedData2['recVer'] && $fontEmbedData2['recInstance'] >= 0x000 && $fontEmbedData2['recInstance'] <= 0x003 && self::RT_FONTEMBEDDATABLOB == $fontEmbedData2['recType']) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData2['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData2['recLen'];
                    }

                    $fontEmbedData3 = $this->loadRecordHeader($stream, $pos);
                    if (0x0 == $fontEmbedData3['recVer'] && $fontEmbedData3['recInstance'] >= 0x000 && $fontEmbedData3['recInstance'] <= 0x003 && self::RT_FONTEMBEDDATABLOB == $fontEmbedData3['recType']) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData3['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData3['recLen'];
                    }

                    $fontEmbedData4 = $this->loadRecordHeader($stream, $pos);
                    if (0x0 == $fontEmbedData4['recVer'] && $fontEmbedData4['recInstance'] >= 0x000 && $fontEmbedData4['recInstance'] <= 0x003 && self::RT_FONTEMBEDDATABLOB == $fontEmbedData4['recType']) {
                        $pos += 8;
                        $fontCollection['recLen'] -= 8;
                        $pos += $fontEmbedData4['recLen'];
                        $fontCollection['recLen'] -= $fontEmbedData4['recLen'];
                    }
                } while ($fontCollection['recLen'] > 0);
            }

            $textCFDefaultsAtom = $this->loadRecordHeader($stream, $pos);
            if (0x0 == $textCFDefaultsAtom['recVer'] && 0x000 == $textCFDefaultsAtom['recInstance'] && self::RT_TEXTCHARFORMATEXCEPTIONATOM == $textCFDefaultsAtom['recType']) {
                $pos += 8;
                $pos += $textCFDefaultsAtom['recLen'];
            }

            $textPFDefaultsAtom = $this->loadRecordHeader($stream, $pos);
            if (0x0 == $textPFDefaultsAtom['recVer'] && 0x000 == $textPFDefaultsAtom['recInstance'] && self::RT_TEXTPARAGRAPHFORMATEXCEPTIONATOM == $textPFDefaultsAtom['recType']) {
                $pos += 8;
                $pos += $textPFDefaultsAtom['recLen'];
            }

            $defaultRulerAtom = $this->loadRecordHeader($stream, $pos);
            if (0x0 == $defaultRulerAtom['recVer'] && 0x000 == $defaultRulerAtom['recInstance'] && self::RT_DEFAULTRULERATOM == $defaultRulerAtom['recType']) {
                $pos += 8;
                $pos += $defaultRulerAtom['recLen'];
            }

            $textSIDefaultsAtom = $this->loadRecordHeader($stream, $pos);
            if (0x0 == $textSIDefaultsAtom['recVer'] && 0x000 == $textSIDefaultsAtom['recInstance'] && self::RT_TEXTSPECIALINFODEFAULTATOM == $textSIDefaultsAtom['recType']) {
                $pos += 8;
                $pos += $textSIDefaultsAtom['recLen'];
            }

            $textMasterStyleAtom = $this->loadRecordHeader($stream, $pos);
            if (0x0 == $textMasterStyleAtom['recVer'] && self::RT_TEXTMASTERSTYLEATOM == $textMasterStyleAtom['recType']) {
                $pos += 8;
                $pos += $textMasterStyleAtom['recLen'];
            }
        }

        $soundCollection = $this->loadRecordHeader($stream, $pos);
        if (0xF == $soundCollection['recVer'] && 0x005 == $soundCollection['recInstance'] && self::RT_SOUNDCOLLECTION == $soundCollection['recType']) {
            $pos += 8;
            $pos += $soundCollection['recLen'];
        }

        $drawingGroup = $this->loadRecordHeader($stream, $pos);
        if (0xF == $drawingGroup['recVer'] && 0x000 == $drawingGroup['recInstance'] && self::RT_DRAWINGGROUP == $drawingGroup['recType']) {
            $drawing = $this->readRecordDrawingGroupContainer($stream, $pos);
            $pos += 8;
            $pos += $drawing['length'];
        }

        $masterList = $this->loadRecordHeader($stream, $pos);
        if (0xF == $masterList['recVer'] && 0x001 == $masterList['recInstance'] && self::RT_SLIDELISTWITHTEXT == $masterList['recType']) {
            $pos += 8;
            $pos += $masterList['recLen'];
        }

        $docInfoList = $this->loadRecordHeader($stream, $pos);
        if (0xF == $docInfoList['recVer'] && 0x000 == $docInfoList['recInstance'] && self::RT_LIST == $docInfoList['recType']) {
            $pos += 8;
            $pos += $docInfoList['recLen'];
        }

        $slideHF = $this->loadRecordHeader($stream, $pos);
        if (0xF == $slideHF['recVer'] && 0x003 == $slideHF['recInstance'] && self::RT_HEADERSFOOTERS == $slideHF['recType']) {
            $pos += 8;
            $pos += $slideHF['recLen'];
        }

        $notesHF = $this->loadRecordHeader($stream, $pos);
        if (0xF == $notesHF['recVer'] && 0x004 == $notesHF['recInstance'] && self::RT_HEADERSFOOTERS == $notesHF['recType']) {
            $pos += 8;
            $pos += $notesHF['recLen'];
        }

        // SlideListWithTextContainer
        $slideList = $this->loadRecordHeader($stream, $pos);
        if (0xF == $slideList['recVer'] && 0x000 == $slideList['recInstance'] && self::RT_SLIDELISTWITHTEXT == $slideList['recType']) {
            $pos += 8;
            do {
                // SlideListWithTextSubContainerOrAtom
                $rhSlideList = $this->loadRecordHeader($stream, $pos);
                if (0x0 == $rhSlideList['recVer'] && 0x000 == $rhSlideList['recInstance'] && self::RT_SLIDEPERSISTATOM == $rhSlideList['recType'] && 0x00000014 == $rhSlideList['recLen']) {
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
                    if (-2147483648 == $slideId) {
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
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd923801(v=office.12).aspx
     */
    private function readRecordDrawingContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_DRAWING == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;

            $officeArtDg = $this->readRecordOfficeArtDgContainer($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $officeArtDg['length'];
        }

        return $arrayReturn;
    }

    /**
     * @return array<string, int>
     */
    private function readRecordDrawingGroupContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_DRAWINGGROUP == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a reference to an external object.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd910388(v=office.12).aspx
     */
    private function readRecordExObjRefAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_EXTERNALOBJECTREFATOM == $data['recType'] && 0x00000004 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a type of action to be performed.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd953300(v=office.12).aspx
     */
    private function readRecordInteractiveInfoAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_INTERACTIVEINFOATOM == $data['recType'] && 0x00000010 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // soundIdRef
            $arrayReturn['length'] += 4;
            // exHyperlinkIdRef
            $arrayReturn['exHyperlinkIdRef'] = self::getInt4d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 4;
            // action
            ++$arrayReturn['length'];
            // oleVerb
            ++$arrayReturn['length'];
            // jump
            ++$arrayReturn['length'];
            // fAnimated (1 bit)
            // fStopSound (1 bit)
            // fCustomShowReturn (1 bit)
            // fVisited (1 bit)
            // reserved (4 bits)
            ++$arrayReturn['length'];
            // hyperlinkType
            ++$arrayReturn['length'];
            // unused
            $arrayReturn['length'] += 3;
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies the name of a macro, a file name, or a named show.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd925121(v=office.12).aspx
     */
    private function readRecordMacroNameAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x002 == $data['recInstance'] && self::RT_CSTRING == $data['recType'] && $data['recLen'] % 2 == 0) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies what actions to perform when interacting with an object by means of a mouse click.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd952348(v=office.12).aspx
     */
    private function readRecordMouseClickInteractiveInfoContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_INTERACTIVEINFO == $data['recType']) {
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
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd925811(v=office.12).aspx
     */
    private function readRecordMouseOverInteractiveInfoContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x001 == $data['recInstance'] && self::RT_INTERACTIVEINFO == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;

            // interactiveInfoAtom
            // macroNameAtom
            throw new FeatureNotImplementedException();
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtBlip record specifies BLIP file data.
     *
     * @return array{'length': int, 'picture': null|string}
     *
     * @see https://msdn.microsoft.com/en-us/library/dd910081(v=office.12).aspx
     */
    private function readRecordOfficeArtBlip(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
            'picture' => null,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && ($data['recType'] >= 0xF018 && $data['recType'] <= 0xF117)) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            switch ($data['recType']) {
                case self::OFFICEARTBLIPJPG:
                case self::OFFICEARTBLIPPNG:
                    // rgbUid1
                    $arrayReturn['length'] += 16;
                    $data['recLen'] -= 16;
                    if (0x6E1 == $data['recInstance']) {
                        // rgbUid2
                        $arrayReturn['length'] += 16;
                        $data['recLen'] -= 16;
                    }
                    // tag
                    ++$arrayReturn['length'];
                    --$data['recLen'];
                    // BLIPFileData
                    $arrayReturn['picture'] = substr($this->streamPictures, $pos + $arrayReturn['length'], $data['recLen']);
                    $arrayReturn['length'] += $data['recLen'];

                    break;
                default:
                    // var_dump(dechex((int) $data['recType']))
                    throw new FeatureNotImplementedException();
            }
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtChildAnchor record specifies four signed integers that specify the anchor for the shape that contains this record.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd922720(v=office.12).aspx
     */
    private function readRecordOfficeArtChildAnchor(string $stream, int $pos)
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF00F == $data['recType'] && 0x00000010 == $data['recLen']) {
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
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd922797(v=office.12).aspx
     */
    private function readRecordOfficeArtClientAnchor(string $stream, int $pos)
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF010 == $data['recType'] && (0x00000008 == $data['recLen'] || 0x00000010 == $data['recLen'])) {
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
                    // record OfficeArtClientAnchor (0x00000010)
                    throw new FeatureNotImplementedException();
            }
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies text related data for a shape.
     *
     * @return array{'length': int, 'alignH': null|string, 'text': string, 'numParts': int, 'numTexts': int, 'hyperlink': array<int, array<string, int>>, 'part': array{'length': int, 'strLenRT': int, 'partLength': float|int, 'bold': bool, 'italic': bool, 'underline': bool, 'fontName': string, 'fontSize': int, 'color': Color}}
     *
     * @see https://msdn.microsoft.com/en-us/library/dd910958(v=office.12).aspx
     */
    private function readRecordOfficeArtClientTextbox(string $stream, int $pos)
    {
        $arrayReturn = [
            'length' => 0,
            'text' => '',
            'numParts' => 0,
            'numTexts' => 0,
            'hyperlink' => [],
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        // recVer 0xF
        // Doc : 0x0    https://msdn.microsoft.com/en-us/library/dd910958(v=office.12).aspx
        // Sample : 0xF https://msdn.microsoft.com/en-us/library/dd953497(v=office.12).aspx
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF00D == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $strLen = 0;
            do {
                $rhChild = $this->loadRecordHeader($stream, $pos + $arrayReturn['length']);
                // @link : https://msdn.microsoft.com/en-us/library/dd947039(v=office.12).aspx
                // echo dechex($rhChild['recType']).'-'.$rhChild['recType'].EOL;
                switch ($rhChild['recType']) {
                    case self::RT_INTERACTIVEINFO:
                        //@link : http://msdn.microsoft.com/en-us/library/dd948623(v=office.12).aspx
                        if (0x0000 == $rhChild['recInstance']) {
                            $mouseClickInfo = $this->readRecordMouseClickInteractiveInfoContainer($stream, $pos + $arrayReturn['length']);
                            $arrayReturn['length'] += $mouseClickInfo['length'];
                            $arrayReturn['hyperlink'][]['id'] = $mouseClickInfo['exHyperlinkIdRef'];
                        }
                        if (0x0001 == $rhChild['recInstance']) {
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
                            ++$arrayReturn['numTexts'];
                            $arrayReturn['text' . $arrayReturn['numTexts']] = $strucTextPFRun;
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
                            ++$arrayReturn['numParts'];
                            $arrayReturn['part' . $arrayReturn['numParts']] = $strucTextCFRun;
                            $strLenRT = $strucTextCFRun['strLenRT'];
                            $arrayReturn['length'] += $strucTextCFRun['length'];
                        } while ($strLenRT > 0);

                        break;
                    case self::RT_TEXTBYTESATOM:
                        $arrayReturn['length'] += 8;
                        // @link : https://msdn.microsoft.com/en-us/library/dd947905(v=office.12).aspx
                        $strLen = (int) $rhChild['recLen'];
                        for ($inc = 0; $inc < $strLen; ++$inc) {
                            $char = self::getInt1d($stream, $pos + $arrayReturn['length']);
                            if (0x0B == $char) {
                                $char = 0x20;
                            }
                            $arrayReturn['text'] .= Text::chr($char);
                            ++$arrayReturn['length'];
                        }

                        break;
                    case self::RT_TEXTCHARSATOM:
                        $arrayReturn['length'] += 8;
                        // @link : http://msdn.microsoft.com/en-us/library/dd772921(v=office.12).aspx
                        $strLen = (int) ($rhChild['recLen'] / 2);
                        for ($inc = 0; $inc < $strLen; ++$inc) {
                            $char = self::getInt2d($stream, $pos + $arrayReturn['length']);
                            if (0x0B == $char) {
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
                        if (0x0000 == $rhChild['recInstance']) {
                            //@todo : MouseClickTextInteractiveInfoAtom
                            $arrayReturn['hyperlink'][count($arrayReturn['hyperlink']) - 1]['start'] = self::getInt4d($stream, $pos + +$arrayReturn['length']);
                            $arrayReturn['length'] += 4;

                            $arrayReturn['hyperlink'][count($arrayReturn['hyperlink']) - 1]['end'] = self::getInt4d($stream, $pos + +$arrayReturn['length']);
                            $arrayReturn['length'] += 4;
                        }
                        if (0x0001 == $rhChild['recInstance']) {
                            throw new FeatureNotImplementedException();
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
                }
            } while (($data['recLen'] - $arrayReturn['length']) > 0);
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtSpContainer record specifies a shape container.
     *
     * @return array{'length': int, 'shape': null|AbstractShape}
     *
     * @see https://msdn.microsoft.com/en-us/library/dd943794(v=office.12).aspx
     */
    private function readRecordOfficeArtSpContainer(string $stream, int $pos)
    {
        $arrayReturn = [
            'length' => 0,
            'shape' => null,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF004 == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // shapeGroup
            $shapeGroup = $this->readRecordOfficeArtFSPGR($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += $shapeGroup['length'];

            // shapeProp
            $shapeProp = $this->readRecordOfficeArtFSP($stream, $pos + $arrayReturn['length']);
            if (0 == $shapeProp['length']) {
                throw new InvalidFileFormatException($this->filename, self::class);
            }
            $arrayReturn['length'] += $shapeProp['length'];

            if (0x1 == $shapeProp['fDeleted'] && 0x0 == $shapeProp['fChild']) {
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
            if (0 == $shpSecondaryOptions1['length']) {
                $shpSecondaryOptions2 = $this->readRecordOfficeArtSecondaryFOPT($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += $shpSecondaryOptions2['length'];
            }

            // shapeTertiaryOptions2
            if (0 == $shpTertiaryOptions1['length']) {
                $shpTertiaryOptions2 = $this->readRecordOfficeArtTertiaryFOPT($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += $shpTertiaryOptions2['length'];
            }

            // Core : Shape
            // Informations about group are not defined
            $arrayDimensions = [];
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
                    if ($this->loadImages && isset($this->arrayPictures[$drawingPib - 1])) {
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
                    // @phpstan-ignore-next-line
                    for ($inc = 1; $inc <= $clientTextbox['numParts']; ++$inc) {
                        if ($clientTextbox['numParts'] == $clientTextbox['numTexts'] && isset($clientTextbox['text' . $inc])) {
                            if (isset($clientTextbox['text' . $inc]['bulletChar'])) {
                                $arrayReturn['shape']->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
                                $arrayReturn['shape']->getActiveParagraph()->getBulletStyle()->setBulletChar($clientTextbox['text' . $inc]['bulletChar']);
                            }
                            // Indent
                            $indent = 0;
                            if (isset($clientTextbox['text' . $inc]['indent'])) {
                                $indent = $clientTextbox['text' . $inc]['indent'];
                            }
                            if (isset($clientTextbox['text' . $inc]['leftMargin'])) {
                                if ($lastMarginLeft > $clientTextbox['text' . $inc]['leftMargin']) {
                                    --$lastLevel;
                                }
                                if ($lastMarginLeft < $clientTextbox['text' . $inc]['leftMargin']) {
                                    ++$lastLevel;
                                }
                                $arrayReturn['shape']->getActiveParagraph()->getAlignment()->setLevel($lastLevel);
                                $lastMarginLeft = $clientTextbox['text' . $inc]['leftMargin'];

                                $arrayReturn['shape']->getActiveParagraph()->getAlignment()->setMarginLeft($clientTextbox['text' . $inc]['leftMargin']);
                                $arrayReturn['shape']->getActiveParagraph()->getAlignment()->setIndent($indent - $clientTextbox['text' . $inc]['leftMargin']);
                            }
                        }
                        // Texte
                        $sText = substr($clientTextbox['text'] ?? '', $start, $clientTextbox['part' . $inc]['partLength']);
                        $sHyperlinkURL = '';
                        if (empty($sText)) {
                            // Is there a hyperlink ?
                            if (!empty($clientTextbox['hyperlink'])) {
                                foreach ($clientTextbox['hyperlink'] as $itmHyperlink) {
                                    if ($itmHyperlink['start'] == $start && ($itmHyperlink['end'] - $itmHyperlink['start']) == (float) $clientTextbox['part' . $inc]['partLength']) {
                                        $sText = $this->arrayHyperlinks[$itmHyperlink['id']]['text'];
                                        $sHyperlinkURL = $this->arrayHyperlinks[$itmHyperlink['id']]['url'];

                                        break;
                                    }
                                }
                            }
                        }
                        // New paragraph
                        $bCreateParagraph = false;
                        if (false !== strpos($sText, "\r")) {
                            $bCreateParagraph = true;
                            $sText = str_replace("\r", '', $sText);
                        }
                        // TextRun
                        $txtRun = $arrayReturn['shape']->createTextRun($sText);
                        if (isset($clientTextbox['part' . $inc]['bold'])) {
                            $txtRun->getFont()->setBold($clientTextbox['part' . $inc]['bold']);
                        }
                        if (isset($clientTextbox['part' . $inc]['italic'])) {
                            $txtRun->getFont()->setItalic($clientTextbox['part' . $inc]['italic']);
                        }
                        if (isset($clientTextbox['part' . $inc]['underline'])) {
                            $txtRun->getFont()->setUnderline(
                                $clientTextbox['part' . $inc]['underline'] ? Font::UNDERLINE_SINGLE : Font::UNDERLINE_NONE
                            );
                        }
                        if (isset($clientTextbox['part' . $inc]['fontName'])) {
                            $txtRun->getFont()->setName($clientTextbox['part' . $inc]['fontName']);
                        }
                        if (isset($clientTextbox['part' . $inc]['fontSize'])) {
                            $txtRun->getFont()->setSize($clientTextbox['part' . $inc]['fontSize']);
                        }
                        if (isset($clientTextbox['part' . $inc]['color'])) {
                            $txtRun->getFont()->setColor($clientTextbox['part' . $inc]['color']);
                        }
                        // Hyperlink
                        if (!empty($sHyperlinkURL)) {
                            $txtRun->setHyperlink(new Hyperlink($sHyperlinkURL));
                        }

                        $start += $clientTextbox['part' . $inc]['partLength'];
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
                    if (isset($shpPrimaryOptions['shadowOffsetX'], $shpPrimaryOptions['shadowOffsetY'])) {
                        $shadowOffsetX = $shpPrimaryOptions['shadowOffsetX'];
                        $shadowOffsetY = $shpPrimaryOptions['shadowOffsetY'];
                        if (0 != $shadowOffsetX && 0 != $shadowOffsetX) {
                            $arrayReturn['shape']->getShadow()->setVisible(true);
                            if ($shadowOffsetX > 0 && $shadowOffsetX == $shadowOffsetY) {
                                $arrayReturn['shape']->getShadow()->setDistance($shadowOffsetX)->setDirection(45);
                            }
                        }
                    }
                    // Specific Line
                    if ($arrayReturn['shape'] instanceof Line) {
                        if (isset($shpPrimaryOptions['lineColor'])) {
                            $arrayReturn['shape']->getBorder()->getColor()->setARGB('FF' . $shpPrimaryOptions['lineColor']);
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
     *
     * @param string $stream
     * @param int $pos
     * @param bool $bInGroup
     *
     * @return array<string, int>
     *
     * @see : https://msdn.microsoft.com/en-us/library/dd910416(v=office.12).aspx
     */
    private function readRecordOfficeArtSpgrContainer($stream, $pos, $bInGroup = false)
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF003 == $data['recType']) {
            $arrayReturn['length'] += 8;

            do {
                $rhFileBlock = $this->loadRecordHeader($stream, $pos + $arrayReturn['length']);
                if (!(0xF == $rhFileBlock['recVer'] && 0x0000 == $rhFileBlock['recInstance'] && (0xF003 == $rhFileBlock['recType'] || 0xF004 == $rhFileBlock['recType']))) {
                    throw new InvalidFileFormatException($this->filename, self::class);
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
                        if (null !== $fileBlock['shape']) {
                            switch ($this->inMainType) {
                                case self::RT_NOTES:
                                    $arrayIdxSlide = array_flip($this->arrayNotes);
                                    if ($this->currentNote > 0 && isset($arrayIdxSlide[$this->currentNote])) {
                                        $oSlide = $this->oPhpPresentation->getSlide($arrayIdxSlide[$this->currentNote]);
                                        if (0 == count($oSlide->getNote()->getShapeCollection())) {
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
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd950206(v=office.12).aspx
     */
    private function readRecordOfficeArtTertiaryFOPT(string $stream, int $pos)
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x3 == $data['recVer'] && 0xF122 == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;

            $officeArtFOPTE = [];
            for ($inc = 0; $inc < $data['recInstance']; ++$inc) {
                $opid = self::getInt2d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $optOp = self::getInt4d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 4;
                $officeArtFOPTE[] = [
                    'opid' => ($opid >> 0) & bindec('11111111111111'),
                    'fBid' => ($opid >> 14) & bindec('1'),
                    'fComplex' => ($opid >> 15) & bindec('1'),
                    'op' => $optOp,
                ];
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
                        if (0x1 == $opt['fComplex']) {
                            $arrayReturn['length'] += $opt['op'];
                        }

                        break;
                    case 0x03A9:
                        // GroupShape : metroBlob
                        //@link : https://msdn.microsoft.com/en-us/library/dd943388(v=office.12).aspx
                        if (0x1 == $opt['fComplex']) {
                            $arrayReturn['length'] += $opt['op'];
                        }

                        break;
                    case 0x01FF:
                        // Line Style Boolean
                        //@link : https://msdn.microsoft.com/en-us/library/dd951605(v=office.12).aspx
                        break;
                    default:
                        // var_dump('0x' . dechex($opt['opid']));
                        throw new FeatureNotImplementedException();
                }
            }
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtDgContainer record specifies the container for all the file records for the objects in a drawing.
     *
     * @return array<string, int>
     *
     * @see : https://msdn.microsoft.com/en-us/library/dd924455(v=office.12).aspx
     */
    private function readRecordOfficeArtDgContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF002 == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // drawingData
            $drawingData = $this->readRecordOfficeArtFDG($stream, $pos + $arrayReturn['length']);
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
     *
     * @return array<string, int>
     *
     * @see : https://msdn.microsoft.com/en-us/library/dd946757(v=office.12).aspx
     */
    private function readRecordOfficeArtFDG(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && $data['recInstance'] <= 0xFFE && 0xF008 == $data['recType'] && 0x00000008 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtFOPT record specifies a table of OfficeArtRGFOPTE records.
     *
     * @return array<string, bool|int|string>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd943404(v=office.12).aspx
     */
    private function readRecordOfficeArtFOPT(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x3 == $data['recVer'] && 0xF00B == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;

            //@link : http://msdn.microsoft.com/en-us/library/dd906086(v=office.12).aspx
            $officeArtFOPTE = [];
            for ($inc = 0; $inc < $data['recInstance']; ++$inc) {
                $opid = self::getInt2d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $data['recLen'] -= 2;
                $optOp = self::getInt4d($this->streamPowerpointDocument, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 4;
                $data['recLen'] -= 4;
                $officeArtFOPTE[] = [
                    'opid' => ($opid >> 0) & bindec('11111111111111'),
                    'fBid' => ($opid >> 14) & bindec('1'),
                    'fComplex' => ($opid >> 15) & bindec('1'),
                    'op' => $optOp,
                ];
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
                        $arrayReturn['insetLeft'] = (int) \PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']);

                        break;
                    case 0x0082:
                        // Text : dyTextTop
                        //@link : http://msdn.microsoft.com/en-us/library/dd925068(v=office.12).aspx
                        $arrayReturn['insetTop'] = (int) \PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']);

                        break;
                    case 0x0083:
                        // Text : dxTextRight
                        //@link : http://msdn.microsoft.com/en-us/library/dd906782(v=office.12).aspx
                        $arrayReturn['insetRight'] = (int) \PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']);

                        break;
                    case 0x0084:
                        // Text : dyTextBottom
                        //@link : http://msdn.microsoft.com/en-us/library/dd772858(v=office.12).aspx
                        $arrayReturn['insetBottom'] = (int) \PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']);

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
                        if (0 == $opt['fComplex']) {
                            $arrayReturn['pib'] = $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }
                        // pib Complex

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
                        if (1 == $opt['fComplex']) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }

                        break;
                    case 0x146:
                        // Geometry : pSegmentInfo
                        //@link : http://msdn.microsoft.com/en-us/library/dd905742(v=office.12).aspx
                        if (1 == $opt['fComplex']) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }

                        break;
                    case 0x155:
                        // Geometry : pAdjustHandles
                        //@link : http://msdn.microsoft.com/en-us/library/dd905890(v=office.12).aspx
                        if (1 == $opt['fComplex']) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }

                        break;
                    case 0x156:
                        // Geometry : pGuides
                        //@link : http://msdn.microsoft.com/en-us/library/dd910801(v=office.12).aspx
                        if (1 == $opt['fComplex']) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }

                        break;
                    case 0x157:
                        // Geometry : pInscribe
                        //@link : http://msdn.microsoft.com/en-us/library/dd904889(v=office.12).aspx
                        if (1 == $opt['fComplex']) {
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
                        $strColor = str_pad(dechex(($opt['op'] >> 0) & bindec('11111111')), 2, '0', STR_PAD_LEFT);
                        $strColor .= str_pad(dechex(($opt['op'] >> 8) & bindec('11111111')), 2, '0', STR_PAD_LEFT);
                        $strColor .= str_pad(dechex(($opt['op'] >> 16) & bindec('11111111')), 2, '0', STR_PAD_LEFT);

                        // echo 'fillColor  : '.$strColor.EOL;
                        break;
                    case 0x0183:
                        // Fill : fillBackColor
                        //@link : http://msdn.microsoft.com/en-us/library/dd950634(v=office.12).aspx
                        $strColor = str_pad(dechex(($opt['op'] >> 0) & bindec('11111111')), 2, '0', STR_PAD_LEFT);
                        $strColor .= str_pad(dechex(($opt['op'] >> 8) & bindec('11111111')), 2, '0', STR_PAD_LEFT);
                        $strColor .= str_pad(dechex(($opt['op'] >> 16) & bindec('11111111')), 2, '0', STR_PAD_LEFT);

                        // echo 'fillBackColor  : '.$strColor.EOL;
                        break;
                    case 0x0193:
                        // Fill : fillRectRight
                        //@link : http://msdn.microsoft.com/en-us/library/dd951294(v=office.12).aspx
                        // echo 'fillRectRight  : '.\PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']).EOL;
                        break;
                    case 0x0194:
                        // Fill : fillRectBottom
                        //@link : http://msdn.microsoft.com/en-us/library/dd910194(v=office.12).aspx
                        // echo 'fillRectBottom   : '.\PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']).EOL;
                        break;
                    case 0x01BF:
                        // Fill : Fill Style Boolean Properties
                        //@link : http://msdn.microsoft.com/en-us/library/dd909380(v=office.12).aspx
                        break;
                    case 0x01C0:
                        // Line Style : lineColor
                        //@link : http://msdn.microsoft.com/en-us/library/dd920397(v=office.12).aspx
                        $strColor = str_pad(dechex(($opt['op'] >> 0) & bindec('11111111')), 2, '0', STR_PAD_LEFT);
                        $strColor .= str_pad(dechex(($opt['op'] >> 8) & bindec('11111111')), 2, '0', STR_PAD_LEFT);
                        $strColor .= str_pad(dechex(($opt['op'] >> 16) & bindec('11111111')), 2, '0', STR_PAD_LEFT);
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
                        $arrayReturn['lineWidth'] = (int) \PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']);

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
                        $arrayReturn['shadowOffsetX'] = (int) \PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']);

                        break;
                    case 0x0206:
                        // Shadow Style : shadowOffsetY
                        //@link : http://msdn.microsoft.com/en-us/library/dd907855(v=office.12).aspx
                        $arrayReturn['shadowOffsetY'] = (int) \PhpOffice\Common\Drawing::emuToPixels((int) $opt['op']);

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
                        if (1 == $opt['fComplex']) {
                            $arrayReturn['length'] += $opt['op'];
                            $data['recLen'] -= $opt['op'];
                        }

                        break;
                    case 0x03BF:
                        // Group Shape Property Set : Group Shape Boolean Properties
                        //@link : http://msdn.microsoft.com/en-us/library/dd949807(v=office.12).aspx
                        break;
                    default:
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
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd947479(v=office.12).aspx
     */
    private function readRecordOfficeArtFPSPL(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF11D == $data['recType'] && 0x00000004 == $data['recLen']) {
            $arrayReturn['length'] += 8;
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * The OfficeArtFSP record specifies an instance of a shape.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd925898(v=office.12).aspx
     */
    private function readRecordOfficeArtFSP(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x2 == $data['recVer'] && 0xF00A == $data['recType'] && 0x00000008 == $data['recLen']) {
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
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd925381(v=office.12).aspx
     */
    private function readRecordOfficeArtFSPGR(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x1 == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF009 == $data['recType'] && 0x00000010 == $data['recLen']) {
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
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd950259(v=office.12).aspx
     */
    private function readRecordOfficeArtSecondaryFOPT(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x3 == $data['recVer'] && 0xF121 == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies information about a shape.
     *
     * @return array<string, int>
     *
     * @see : https://msdn.microsoft.com/en-us/library/dd950927(v=office.12).aspx
     */
    private function readRecordOfficeArtClientData(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && 0xF011 == $data['recType']) {
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
            $array = [
                self::RT_PROGTAGS,
                self::RT_ROUNDTRIPNEWPLACEHOLDERID12ATOM,
                self::RT_ROUNDTRIPSHAPEID12ATOM,
                self::RT_ROUNDTRIPHFPLACEHOLDER12ATOM,
                self::RT_ROUNDTRIPSHAPECHECKSUMFORCL12ATOM,
            ];
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
                            // var_dump('0x' . dechex($dataHeaderRG['recType']));
                            throw new FeatureNotImplementedException();
                    }
                }
            } while (in_array($dataHeaderRG['recType'], $array));
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a persist object directory. Each persist object identifier specified MUST be unique in that persist object directory.
     *
     * @see http://msdn.microsoft.com/en-us/library/dd952680(v=office.12).aspx
     */
    private function readRecordPersistDirectoryAtom(string $stream, int $pos): void
    {
        $rHeader = $this->loadRecordHeader($stream, $pos);
        $pos += 8;
        if (0x0 != $rHeader['recVer'] || 0x000 != $rHeader['recInstance'] || self::RT_PERSISTDIRECTORYATOM != $rHeader['recType']) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : PersistDirectoryAtom > RecordHeader');
        }
        // rgPersistDirEntry
        // @link : http://msdn.microsoft.com/en-us/library/dd947347(v=office.12).aspx
        do {
            $data = self::getInt4d($stream, $pos);
            $pos += 4;
            $rHeader['recLen'] -= 4;
            //$persistId  = ($data >> 0) & bindec('11111111111111111111');
            $cPersist = ($data >> 20) & bindec('111111111111');

            $rgPersistOffset = [];
            for ($inc = 0; $inc < $cPersist; ++$inc) {
                $rgPersistOffset[] = self::getInt4d($stream, $pos);
                $pos += 4;
                $rHeader['recLen'] -= 4;
            }
        } while ($rHeader['recLen'] > 0);
        $this->rgPersistDirEntry = $rgPersistOffset;
    }

    /**
     * A container record that specifies information about the headers (1) and footers within a slide.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd904856(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordPerSlideHeadersFootersContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_HEADERSFOOTERS == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies whether a shape is a placeholder shape.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd923930(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordPlaceholderAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_PLACEHOLDERATOM == $data['recType'] && 0x00000008 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a collection of re-color mappings for a metafile ([MS-WMF]).
     *
     * @see https://msdn.microsoft.com/en-us/library/dd904899(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordRecolorInfoAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_RECOLORINFOATOM == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies that a shape is a header or footerplaceholder shape.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd910800(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordRoundTripHFPlaceholder12Atom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_ROUNDTRIPHFPLACEHOLDER12ATOM == $data['recType'] && 0x00000001 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a shape identifier.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd772926(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordRoundTripShapeId12Atom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_ROUNDTRIPSHAPEID12ATOM == $data['recType'] && 0x00000004 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies information about a slide that synchronizes to a slide in a slide library.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd923801(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordRoundTripSlideSyncInfo12Container(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_ROUNDTRIPSLIDESYNCINFO12 == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies shape-level Boolean flags.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd908949(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordShapeFlags10Atom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_SHAPEFLAGS10ATOM == $data['recType'] && 0x00000001 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies shape-level Boolean flags.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd925824(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordShapeFlagsAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_SHAPEATOM == $data['recType'] && 0x00000001 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies programmable tags with additional binary shape data.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd911033(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordShapeProgBinaryTagContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_PROGBINARYTAG == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies programmable tags with additional shape data.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd911266(v=office.12).aspx
     */
    private function readRecordShapeProgTagsContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_PROGTAGS == $data['recType']) {
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
                        // var_dump('0x' . dechex($dataHeaderRG['recType']));
                        throw new FeatureNotImplementedException();
                }
            } while ($length < $data['recLen']);
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies information about a slide.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd923801(v=office.12).aspx
     */
    private function readRecordSlideAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x2 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_SLIDEATOM == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // slideAtom > geom
            $arrayReturn['length'] += 4;
            // slideAtom > rgPlaceholderTypes
            $rgPlaceholderTypes = [];
            for ($inc = 0; $inc < 8; ++$inc) {
                $rgPlaceholderTypes[] = self::getInt1d($this->streamPowerpointDocument, $pos);
                ++$arrayReturn['length'];
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
     *
     * @see http://msdn.microsoft.com/en-us/library/dd946323(v=office.12).aspx
     */
    private function readRecordSlideContainer(string $stream, int $pos): void
    {
        // Core
        $this->oPhpPresentation->createSlide();
        $this->oPhpPresentation->setActiveSlideIndex($this->oPhpPresentation->getSlideCount() - 1);

        // *** slideAtom (32 bytes)
        $slideAtom = $this->readRecordSlideAtom($stream, $pos);
        if (0 == $slideAtom['length']) {
            throw new InvalidFileFormatException($this->filename, self::class);
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
        if (0 == $slideSchemeColorAtom['length']) {
            // Record SlideSchemeColorSchemeAtom
            throw new InvalidFileFormatException($this->filename, self::class);
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
     *
     * @see https://msdn.microsoft.com/en-us/library/dd906297(v=office.12).aspx
     *
     * @return array{'length': int, 'slideName': string}
     */
    private function readRecordSlideNameAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
            'slideName' => '',
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x003 == $data['recInstance'] && self::RT_CSTRING == $data['recType'] && $data['recLen'] % 2 == 0) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $strLen = ($data['recLen'] / 2);
            for ($inc = 0; $inc < $strLen; ++$inc) {
                $char = self::getInt2d($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $arrayReturn['slideName'] .= Text::chr($char);
            }
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies a slide number metacharacter.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd945703(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordSlideNumberMCAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_SLIDENUMBERMETACHARATOM == $data['recType'] && 0x00000004 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Datas
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies programmable tags with additional slide data.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd951946(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordSlideProgTagsContainer(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0xF == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_PROGTAGS == $data['recType']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * A container record that specifies the color scheme used by a slide.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd949420(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordSlideSchemeColorSchemeAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x001 == $data['recInstance'] && self::RT_COLORSCHEMEATOM == $data['recType'] && 0x00000020 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length
            $rgSchemeColor = [];
            for ($inc = 0; $inc <= 7; ++$inc) {
                $rgSchemeColor[] = [
                    'red' => self::getInt1d($stream, $pos + $arrayReturn['length'] + $inc * 4),
                    'green' => self::getInt1d($stream, $pos + $arrayReturn['length'] + $inc * 4 + 1),
                    'blue' => self::getInt1d($stream, $pos + $arrayReturn['length'] + $inc * 4 + 2),
                ];
            }
            $arrayReturn['length'] += (8 * 4);
        }

        return $arrayReturn;
    }

    /**
     * An atom record that specifies what transition effect to perform during a slide show, and how to advance to the next presentation slide.
     *
     * @see https://msdn.microsoft.com/en-us/library/dd943408(v=office.12).aspx
     *
     * @return array<string, int>
     */
    private function readRecordSlideShowSlideInfoAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x0 == $data['recVer'] && 0x000 == $data['recInstance'] && self::RT_SLIDESHOWSLIDEINFOATOM == $data['recType'] && 0x00000010 == $data['recLen']) {
            // Record Header
            $arrayReturn['length'] += 8;
            // Length;
            $arrayReturn['length'] += $data['recLen'];
        }

        return $arrayReturn;
    }

    /**
     * UserEditAtom.
     *
     * @see http://msdn.microsoft.com/en-us/library/dd945746(v=office.12).aspx
     */
    private function readRecordUserEditAtom(string $stream, int $pos): void
    {
        $rHeader = $this->loadRecordHeader($stream, $pos);
        $pos += 8;
        if (0x0 != $rHeader['recVer'] || 0x000 != $rHeader['recInstance'] || self::RT_USEREDITATOM != $rHeader['recType'] || (0x0000001C != $rHeader['recLen'] && 0x00000020 != $rHeader['recLen'])) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : UserEditAtom > RecordHeader');
        }

        // lastSlideIdRef
        $pos += 4;
        // version
        $pos += 2;

        // minorVersion
        $minorVersion = self::getInt1d($stream, $pos);
        ++$pos;
        if (0x00 != $minorVersion) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : UserEditAtom > minorVersion');
        }

        // majorVersion
        $majorVersion = self::getInt1d($stream, $pos);
        ++$pos;
        if (0x03 != $majorVersion) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : UserEditAtom > majorVersion');
        }

        // offsetLastEdit
        $pos += 4;
        // offsetPersistDirectory
        $this->offsetPersistDirectory = self::getInt4d($stream, $pos);
        $pos += 4;

        // docPersistIdRef
        $docPersistIdRef = self::getInt4d($stream, $pos);
        $pos += 4;
        if (0x00000001 != $docPersistIdRef) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : UserEditAtom > docPersistIdRef');
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
     *
     * @return array{'length': int, 'strLenRT': int, 'partLength': int, 'bold': bool, 'italic': bool, 'underline': bool, 'fontName': string, 'fontSize': int, 'color': Color}
     *
     * @see https://msdn.microsoft.com/en-us/library/dd945870(v=office.12).aspx
     */
    private function readStructureTextCFRun(string $stream, int $pos, int $strLenRT): array
    {
        $arrayReturn = [
            'length' => 0,
            'strLenRT' => $strLenRT,
        ];

        // rgTextCFRun
        $countRgTextCFRun = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['strLenRT'] -= $countRgTextCFRun;
        $arrayReturn['length'] += 4;
        $arrayReturn['partLength'] = $countRgTextCFRun;

        $masks = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;

        $masksData = [];
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
        if (1 == $masksData['bold'] || 1 == $masksData['italic'] || 1 == $masksData['underline'] || 1 == $masksData['shadow'] || 1 == $masksData['fehint'] || 1 == $masksData['kumi'] || 1 == $masksData['emboss'] || 1 == $masksData['fHasStyle']) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;

            $fontStyleFlags = [];
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

            $arrayReturn['bold'] = (1 == $fontStyleFlags['bold']) ? true : false;
            $arrayReturn['italic'] = (1 == $fontStyleFlags['italic']) ? true : false;
            $arrayReturn['underline'] = (1 == $fontStyleFlags['underline']) ? true : false;
        }
        if (1 == $masksData['typeface']) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayReturn['fontName'] = $this->arrayFonts[$data] ?? '';
        }
        if (1 == $masksData['oldEATypeface']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['ansiTypeface']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['symbolTypeface']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['size']) {
            $arrayReturn['fontSize'] = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['color']) {
            $red = self::getInt1d($stream, $pos + $arrayReturn['length']);
            ++$arrayReturn['length'];
            $green = self::getInt1d($stream, $pos + $arrayReturn['length']);
            ++$arrayReturn['length'];
            $blue = self::getInt1d($stream, $pos + $arrayReturn['length']);
            ++$arrayReturn['length'];
            $index = self::getInt1d($stream, $pos + $arrayReturn['length']);
            ++$arrayReturn['length'];

            if (0xFE == $index) {
                $strColor = str_pad(dechex($red), 2, '0', STR_PAD_LEFT);
                $strColor .= str_pad(dechex($green), 2, '0', STR_PAD_LEFT);
                $strColor .= str_pad(dechex($blue), 2, '0', STR_PAD_LEFT);

                $arrayReturn['color'] = new Color('FF' . $strColor);
            }
        }
        if (1 == $masksData['position']) {
            throw new FeatureNotImplementedException();
        }

        return $arrayReturn;
    }

    /**
     * A structure that specifies the paragraph-level formatting of a run of text.
     *
     * @return array{'length': int, 'strLenRT': int, 'alignH': null|string, 'bulletChar': string, 'leftMargin': int, 'indent': int}
     *
     * @see https://msdn.microsoft.com/en-us/library/dd923535(v=office.12).aspx
     */
    private function readStructureTextPFRun(string $stream, int $pos, int $strLenRT): array
    {
        $arrayReturn = [
            'length' => 0,
            'strLenRT' => $strLenRT,
        ];

        // rgTextPFRun
        $countRgTextPFRun = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['strLenRT'] -= $countRgTextPFRun;
        $arrayReturn['length'] += 4;

        // indent
        $arrayReturn['length'] += 2;

        $masks = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;

        $masksData = [];
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

        $bulletFlags = [];
        if (1 == $masksData['hasBullet'] || 1 == $masksData['bulletHasFont'] || 1 == $masksData['bulletHasColor'] || 1 == $masksData['bulletHasSize']) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;

            $bulletFlags['fHasBullet'] = ($data >> 0) & bindec('1');
            $bulletFlags['fBulletHasFont'] = ($data >> 1) & bindec('1');
            $bulletFlags['fBulletHasColor'] = ($data >> 2) & bindec('1');
            $bulletFlags['fBulletHasSize'] = ($data >> 3) & bindec('1');
        }
        if (1 == $masksData['bulletChar']) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayReturn['bulletChar'] = chr($data);
        }
        if (1 == $masksData['bulletFont']) {
            // $data = self::getInt2d($stream, $pos);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['bulletSize']) {
            // $data = self::getInt2d($stream, $pos);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['bulletColor']) {
            // $red = self::getInt1d($stream, $pos + $arrayReturn['length']);
            ++$arrayReturn['length'];
            // $green = self::getInt1d($stream, $pos + $arrayReturn['length']);
            ++$arrayReturn['length'];
            // $blue = self::getInt1d($stream, $pos + $arrayReturn['length']);
            ++$arrayReturn['length'];
            $index = self::getInt1d($stream, $pos + $arrayReturn['length']);
            ++$arrayReturn['length'];

            if (0xFE == $index) {
                // $strColor = str_pad(dechex($red), 2, '0', STR_PAD_LEFT);
                // $strColor .= str_pad(dechex($green), 2, '0', STR_PAD_LEFT);
                // $strColor .= str_pad(dechex($blue), 2, '0', STR_PAD_LEFT);
            }
        }
        if (1 == $masksData['align']) {
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
        if (1 == $masksData['lineSpacing']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['spaceBefore']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['spaceAfter']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['leftMargin']) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayReturn['leftMargin'] = (int) round($data / 6);
        }
        if (1 == $masksData['indent']) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayReturn['indent'] = (int) round($data / 6);
        }
        if (1 == $masksData['defaultTabSize']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['tabStops']) {
            throw new FeatureNotImplementedException();
        }
        if (1 == $masksData['fontAlign']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['charWrap'] || 1 == $masksData['wordWrap'] || 1 == $masksData['overflow']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['textDirection']) {
            throw new FeatureNotImplementedException();
        }

        return $arrayReturn;
    }

    /**
     * A structure that specifies language and spelling information for a run of text.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd909603(v=office.12).aspx
     */
    private function readStructureTextSIRun(string $stream, int $pos, int $strLenRT): array
    {
        $arrayReturn = [
            'length' => 0,
            'strLenRT' => $strLenRT,
        ];

        $arrayReturn['strLenRT'] -= self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;

        $data = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;
        $masksData = [];
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

        if (1 == $masksData['spell']) {
            $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $masksSpell = [];
            $masksSpell['error'] = ($data >> 0) & bindec('1');
            $masksSpell['clean'] = ($data >> 1) & bindec('1');
            $masksSpell['grammar'] = ($data >> 2) & bindec('1');
        }
        if (1 == $masksData['lang']) {
            // $data = self::getInt2d($stream, $pos);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['altLang']) {
            // $data = self::getInt2d($stream, $pos);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fBidi']) {
            throw new FeatureNotImplementedException();
        }
        if (1 == $masksData['fPp10ext']) {
            throw new FeatureNotImplementedException();
        }
        if (1 == $masksData['smartTag']) {
            throw new FeatureNotImplementedException();
        }

        return $arrayReturn;
    }

    /**
     * A structure that specifies tabbing, margins, and indentation for text.
     *
     * @return array<string, int>
     *
     * @see https://msdn.microsoft.com/en-us/library/dd922749(v=office.12).aspx
     */
    private function readStructureTextRuler(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = self::getInt4d($stream, $pos + $arrayReturn['length']);
        $arrayReturn['length'] += 4;

        $masksData = [];
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

        if (1 == $masksData['fCLevels']) {
            throw new FeatureNotImplementedException();
        }
        if (1 == $masksData['fDefaultTabSize']) {
            throw new FeatureNotImplementedException();
        }
        if (1 == $masksData['fTabStops']) {
            $count = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
            $arrayTabStops = [];
            for ($inc = 0; $inc < $count; ++$inc) {
                $position = self::getInt2d($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $type = self::getInt2d($stream, $pos + $arrayReturn['length']);
                $arrayReturn['length'] += 2;
                $arrayTabStops[] = [
                    'position' => $position,
                    'type' => $type,
                ];
            }
        }
        if (1 == $masksData['fLeftMargin1']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fIndent1']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fLeftMargin2']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fIndent2']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fLeftMargin3']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fIndent3']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fLeftMargin4']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fIndent4']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fLeftMargin5']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }
        if (1 == $masksData['fIndent5']) {
            // $data = self::getInt2d($stream, $pos + $arrayReturn['length']);
            $arrayReturn['length'] += 2;
        }

        return $arrayReturn;
    }

    private function readRecordNotesContainer(string $stream, int $pos): void
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
     * @return array<string, int>
     */
    private function readRecordNotesAtom(string $stream, int $pos): array
    {
        $arrayReturn = [
            'length' => 0,
        ];

        $data = $this->loadRecordHeader($stream, $pos);
        if (0x1 != $data['recVer'] || 0x000 != $data['recInstance'] || self::RT_NOTESATOM != $data['recType'] || 0x00000008 != $data['recLen']) {
            throw new InvalidFileFormatException($this->filename, self::class, 'Location : NotesAtom > RecordHeader)');
        }
        // Record Header
        $arrayReturn['length'] += 8;
        // NotesAtom > slideIdRef
        $notesIdRef = self::getInt4d($stream, $pos + $arrayReturn['length']);
        if (-2147483648 == $notesIdRef) {
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
