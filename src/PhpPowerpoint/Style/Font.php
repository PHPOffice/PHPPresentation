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

namespace PhpOffice\PhpPowerpoint\Style;

use PhpOffice\PhpPowerpoint\ComparableInterface;

/**
 * \PhpOffice\PhpPowerpoint\Style\Font
 */
class Font implements ComparableInterface
{
    /* Underline types */
    const UNDERLINE_NONE = 'none';
    const UNDERLINE_DASH = 'dash';
    const UNDERLINE_DASHHEAVY = 'dashHeavy';
    const UNDERLINE_DASHLONG = 'dashLong';
    const UNDERLINE_DASHLONGHEAVY = 'dashLongHeavy';
    const UNDERLINE_DOUBLE = 'dbl';
    const UNDERLINE_DOTHASH = 'dotDash';
    const UNDERLINE_DOTHASHHEAVY = 'dotDashHeavy';
    const UNDERLINE_DOTDOTDASH = 'dotDotDash';
    const UNDERLINE_DOTDOTDASHHEAVY = 'dotDotDashHeavy';
    const UNDERLINE_DOTTED = 'dotted';
    const UNDERLINE_DOTTEDHEAVY = 'dottedHeavy';
    const UNDERLINE_HEAVY = 'heavy';
    const UNDERLINE_SINGLE = 'sng';
    const UNDERLINE_WAVY = 'wavy';
    const UNDERLINE_WAVYDOUBLE = 'wavyDbl';
    const UNDERLINE_WAVYHEAVY = 'wavyHeavy';
    const UNDERLINE_WORDS = 'words';

    /**
     * Name
     *
     * @var string
     */
    private $name;
    
    /**
     * Font Size
     *
     * @var float|int
     */
    private $size;
    
    /**
     * Bold
     *
     * @var boolean
     */
    private $bold;

    /**
     * Italic
     *
     * @var boolean
     */
    private $italic;

    /**
     * Superscript
     *
     * @var boolean
     */
    private $superScript;

    /**
     * Subscript
     *
     * @var boolean
     */
    private $subScript;

    /**
     * Underline
     *
     * @var string
     */
    private $underline;

    /**
     * Strikethrough
     *
     * @var boolean
     */
    private $strikethrough;

    /**
     * Foreground color
     *
     * @var \PhpOffice\PhpPowerpoint\Style\Color
     */
    private $color;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function __construct()
    {
        // Initialise values
        $this->name          = 'Calibri';
        $this->size          = 10;
        $this->bold          = false;
        $this->italic        = false;
        $this->superScript   = false;
        $this->subScript     = false;
        $this->underline     = self::UNDERLINE_NONE;
        $this->strikethrough = false;
        $this->color         = new Color(Color::COLOR_BLACK);
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Set Name
     *
     * @param  string                   $pValue
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setName($pValue = 'Calibri')
    {
        if ($pValue == '') {
            $pValue = 'Calibri';
        }
        $this->name = $pValue;

        return $this;
    }

    /**
     * Get Size
     *
     * @return double
     */
    public function getSize()
    {
        return $this->size;
    }

    /**
     * Set Size
     *
     * @param float|int $pValue
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setSize($pValue = 10)
    {
        if ($pValue == '') {
            $pValue = 10;
        }
        $this->size = $pValue;

        return $this;
    }

    /**
     * Get Bold
     *
     * @return boolean
     */
    public function isBold()
    {
        return $this->bold;
    }

    /**
     * Set Bold
     *
     * @param  boolean                  $pValue
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setBold($pValue = false)
    {
        if ($pValue == '') {
            $pValue = false;
        }
        $this->bold = $pValue;

        return $this;
    }

    /**
     * Get Italic
     *
     * @return boolean
     */
    public function isItalic()
    {
        return $this->italic;
    }

    /**
     * Set Italic
     *
     * @param  boolean                  $pValue
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setItalic($pValue = false)
    {
        if ($pValue == '') {
            $pValue = false;
        }
        $this->italic = $pValue;

        return $this;
    }

    /**
     * Get SuperScript
     *
     * @return boolean
     */
    public function isSuperScript()
    {
        return $this->superScript;
    }

    /**
     * Set SuperScript
     *
     * @param  boolean                  $pValue
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setSuperScript($pValue = false)
    {
        if ($pValue == '') {
            $pValue = false;
        }
        $this->superScript = $pValue;
        $this->subScript   = !$pValue;

        return $this;
    }

    /**
     * Get SubScript
     *
     * @return boolean
     */
    public function isSubScript()
    {
        return $this->subScript;
    }

    /**
     * Set SubScript
     *
     * @param  boolean                  $pValue
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setSubScript($pValue = false)
    {
        if ($pValue == '') {
            $pValue = false;
        }
        $this->subScript   = $pValue;
        $this->superScript = !$pValue;

        return $this;
    }

    /**
     * Get Underline
     *
     * @return string
     */
    public function getUnderline()
    {
        return $this->underline;
    }

    /**
     * Set Underline
     *
     * @param  string                   $pValue \PhpOffice\PhpPowerpoint\Style\Font underline type
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setUnderline($pValue = self::UNDERLINE_NONE)
    {
        if ($pValue == '') {
            $pValue = self::UNDERLINE_NONE;
        }
        $this->underline = $pValue;

        return $this;
    }

    /**
     * Set Striketrough
     *
     * @deprecated Use setStrikethrough() instead.
     * @param  boolean                  $pValue
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setStriketrough($pValue = false)
    {
        return $this->setStrikethrough($pValue);
    }

    /**
     * Get Strikethrough
     *
     * @return boolean
     */
    public function isStrikethrough()
    {
        return $this->strikethrough;
    }

    /**
     * Set Strikethrough
     *
     * @param  boolean                  $pValue
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setStrikethrough($pValue = false)
    {
        if ($pValue == '') {
            $pValue = false;
        }
        $this->strikethrough = $pValue;

        return $this;
    }

    /**
     * Get Color
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Color
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * Set Color
     *
     * @param  \PhpOffice\PhpPowerpoint\Style\Color $pValue
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function setColor(Color $pValue = null)
    {
        $this->color = $pValue;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->name . $this->size . ($this->bold ? 't' : 'f') . ($this->italic ? 't' : 'f') . ($this->superScript ? 't' : 'f') . ($this->subScript ? 't' : 'f') . $this->underline . ($this->strikethrough ? 't' : 'f') . $this->color->getHashCode() . __CLASS__);
    }

    /**
     * Get hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return string Hash index
     */
    public function getHashIndex()
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param string $value Hash index
     */
    public function setHashIndex($value)
    {
        $this->hashIndex = $value;
    }
}
