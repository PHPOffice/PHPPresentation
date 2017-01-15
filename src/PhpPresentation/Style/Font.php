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

namespace PhpOffice\PhpPresentation\Style;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\Style\Font
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
     * @var \PhpOffice\PhpPresentation\Style\Color
     */
    private $color;

    /**
     * Character Spacing
     *
     * @var int
     */
    private $characterSpacing;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Font
     */
    public function __construct()
    {
        // Initialise values
        $this->name             = 'Calibri';
        $this->size             = 10;
        $this->characterSpacing = 0;
        $this->bold             = false;
        $this->italic           = false;
        $this->superScript      = false;
        $this->subScript        = false;
        $this->underline        = self::UNDERLINE_NONE;
        $this->strikethrough    = false;
        $this->color            = new Color(Color::COLOR_BLACK);
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
     * @return \PhpOffice\PhpPresentation\Style\Font
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
     * Get Character Spacing
     *
     * @return double
     */
    public function getCharacterSpacing()
    {
        return $this->characterSpacing;
    }
    
    /**
     * Set Character Spacing
     * Value in pt
     * @param float|int $pValue
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setCharacterSpacing($pValue = 0)
    {
        if ($pValue == '') {
            $pValue = 0;
        }
        $this->characterSpacing = $pValue * 100;
    
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
     * @return \PhpOffice\PhpPresentation\Style\Font
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
     * @return \PhpOffice\PhpPresentation\Style\Font
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
     * @return \PhpOffice\PhpPresentation\Style\Font
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
     * @return \PhpOffice\PhpPresentation\Style\Font
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
     * @return \PhpOffice\PhpPresentation\Style\Font
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
     * @param  string                   $pValue \PhpOffice\PhpPresentation\Style\Font underline type
     * @return \PhpOffice\PhpPresentation\Style\Font
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
     * @return \PhpOffice\PhpPresentation\Style\Font
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
     * @return \PhpOffice\PhpPresentation\Style\Color|\PhpOffice\PhpPresentation\Style\SchemeColor
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * Set Color
     *
     * @param  \PhpOffice\PhpPresentation\Style\Color|\PhpOffice\PhpPresentation\Style\SchemeColor $pValue
     * @throws \Exception
     * @return \PhpOffice\PhpPresentation\Style\Font
     */
    public function setColor($pValue = null)
    {
        if (!$pValue instanceof Color) {
            throw new \Exception('$pValue must be an instance of \PhpOffice\PhpPresentation\Style\Color');
        }
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
