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
 * \PhpOffice\PhpPresentation\Shape\Effect
 */
class Effect implements ComparableInterface
{
    /** Effect type */
    const EFFECT_SHADOW_INNER = 'innerShdw';
    const EFFECT_SHADOW_OUTER = 'outerShdw';
    const EFFECT_REFLECTION   = 'reflection';
    
    /* Shadow alignment */
    const SHADOW_BOTTOM = 'b';
    const SHADOW_BOTTOM_LEFT = 'bl';
    const SHADOW_BOTTOM_RIGHT = 'br';
    const SHADOW_CENTER = 'ctr';
    const SHADOW_LEFT = 'l';
    const SHADOW_TOP = 't';
    const SHADOW_TOP_LEFT = 'tl';
    const SHADOW_TOP_RIGHT = 'tr';
   
    /**
     * Effect type (inner shadow, outer shadow, reflextion, etc...)
     *
     * @var string
     */
    private ?string $effectType = null;
    /**
     * Effect direction
     *
     * @var int
     */
    private int $direction;
    
    /**
     * Effect distance
     * 
     * @var int
     */
    private int $distance;

    /**
     * Blur radius effect
     * 
     * @var int
     */
    private int $blurRadius;
    
    /**
     * Alignment effect
     * 
     * @var string
     */
    private string $alignment;
    
    /**
     * Color
     *
     * @var \PhpOffice\PhpPresentation\Style\Color
     */
    private $color;

    /**
     * Alpha
     *
     * @var int
     */
    private $alpha;

    /**
     * Hash index
     *
     * @var string
     */
    private string $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Effect instance
     */
    public function __construct(?string $type)
    {
        // Initialise variables
        $this->effectType = $type;
        $this->blurRadius = 6;
        $this->distance   = 2;
        $this->direction  = 0;
        $this->alignment  = self::SHADOW_BOTTOM_RIGHT;
        $this->color      = new Color(Color::COLOR_BLACK);
        $this->alpha      = 50;
    }
    
    /**
     * Define the type effect
     * 
     * @param string $type
     * @return $this
     * @see self::EFFECT_SHADOW_INNER, self::EFFECT_SHADOW_OUTER, self::EFFECT_REFLECTION
     */
    public function setEffectType(string $type)
    {
      $this->effectType = $type;
      return $this;
    }
    
    /**
     * Get the effect type
     * 
     * @return string
     */
    public function getEffectType():string
    {
      return $this->effectType;
    }
    
    /**
     * Set the direction
     *
     * @param int $dir
     * @return $this
     */
    public function setDirection(?int $dir)
    {
      if (!isset($dir)) $dir = 0;
      $this->direction = (int)$dir;
      return $this;
    }
    
    /**
     * Get the direction
     *
     * @return int
     */
    public function getDirection():int
    {
      return $this->direction;
    }
    
    /**
     * Set the blur radius
     * 
     * @param int $radius
     * @return $this
     */
    public function setBlurRadius(?int $radius)
    {
      if (!isset($radius)) $radius = 6;
      $this->blurRadius = $radius;
      return $this;
    }
    
    /**
     * Get the blur radius
     *
     * @return int
     */
    public function getBlurRadius():int
    {
      return $this->blurRadius;
    }
    
    /**
     * Get Shadow distance
     *
     * @return int
     */
    public function getDistance():int
    {
        return $this->distance;
    }

    /**
     * Set Shadow distance
     *
     * @param  int $distance
     * @return $this
     */
    public function setDistance(int $distance = 2)
    {
        $this->distance = $distance;

        return $this;
    }

    /**
     * Set the effect alignment
     * 
     * @param string $align
     * @return $this
     */
    public function setAlignment(?string $align)
    {
      if (!isset($align)) $align = self::SHADOW_BOTTOM_RIGHT;
      $this->align = $align;
      return $this;
    }
    
    /**
     * Get the effect alignment
     * 
     * @return string
     */
    public function getAlignment():string
    {
      return $this->align;
    }

    /**
     * Set Color
     *
     * @param  \PhpOffice\PhpPresentation\Style\Color  $color
     * @return $this
     */
    public function setColor(Color $color = null)
    {
        $this->color = $color;

        return $this;
    }

    /**
     * Get Color
     *
     * @return \PhpOffice\PhpPresentation\Style\Color
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * Set Alpha
     *
     * @param  int $alpha
     * @return $this
     */
    public function setAlpha(?int $alpha)
    {
        if (!isset($alpha)) $alpha = 0;
        $this->alpha = $alpha;

        return $this;
    }

    /**
     * Get Alpha
     *
     * @return int
     */
    public function getAlpha():int
    {
        return $this->alpha;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->effectType . $this->blurRadius . $this->distance . $this->direction . $this->alignment . $this->color->getHashCode() . $this->alpha . __CLASS__);
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
