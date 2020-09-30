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
 * \PhpOffice\PhpPresentation\Style\Shadow
 */
class Shadow implements ComparableInterface
{
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
     * Visible
     *
     * @var boolean
     */
    private $visible;

    /**
     * Blur radius
     *
     * Defaults to 6
     *
     * @var int
     */
    private $blurRadius;

    /**
     * Shadow distance
     *
     * Defaults to 2
     *
     * @var int
     */
    private $distance;

    /**
     * Shadow direction (in degrees)
     *
     * @var int
     */
    private $direction;

    /**
     * Shadow alignment
     *
     * @var string
     */
    private $alignment;

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
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPresentation\Style\Shadow
     */
    public function __construct()
    {
        // Initialise values
        $this->visible    = false;
        $this->blurRadius = 6;
        $this->distance   = 2;
        $this->direction  = 0;
        $this->alignment  = self::SHADOW_BOTTOM_RIGHT;
        $this->color      = new Color(Color::COLOR_BLACK);
        $this->alpha      = 50;
    }

    /**
     * Get Visible
     *
     * @return boolean
     */
    public function isVisible()
    {
        return $this->visible;
    }

    /**
     * Set Visible
     *
     * @param  boolean                    $pValue
     * @return $this
     */
    public function setVisible($pValue = false)
    {
        $this->visible = $pValue;

        return $this;
    }

    /**
     * Get Blur radius
     *
     * @return int
     */
    public function getBlurRadius()
    {
        return $this->blurRadius;
    }

    /**
     * Set Blur radius
     *
     * @param  int                        $pValue
     * @return $this
     */
    public function setBlurRadius($pValue = 6)
    {
        $this->blurRadius = $pValue;

        return $this;
    }

    /**
     * Get Shadow distance
     *
     * @return int
     */
    public function getDistance()
    {
        return $this->distance;
    }

    /**
     * Set Shadow distance
     *
     * @param  int                        $pValue
     * @return $this
     */
    public function setDistance($pValue = 2)
    {
        $this->distance = $pValue;

        return $this;
    }

    /**
     * Get Shadow direction (in degrees)
     *
     * @return int
     */
    public function getDirection()
    {
        return $this->direction;
    }

    /**
     * Set Shadow direction (in degrees)
     *
     * @param  int                        $pValue
     * @return $this
     */
    public function setDirection($pValue = 0)
    {
        $this->direction = $pValue;

        return $this;
    }

    /**
     * Get Shadow alignment
     *
     * @return int
     */
    public function getAlignment()
    {
        return $this->alignment;
    }

    /**
     * Set Shadow alignment
     *
     * @param  string                        $pValue
     * @return $this
     */
    public function setAlignment($pValue = self::SHADOW_BOTTOM_RIGHT)
    {
        $this->alignment = $pValue;

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
     * Set Color
     *
     * @param  \PhpOffice\PhpPresentation\Style\Color  $pValue
     * @throws \Exception
     * @return $this
     */
    public function setColor(Color $pValue = null)
    {
        $this->color = $pValue;

        return $this;
    }

    /**
     * Get Alpha
     *
     * @return int
     */
    public function getAlpha()
    {
        return $this->alpha;
    }

    /**
     * Set Alpha
     *
     * @param  int                        $pValue
     * @return $this
     */
    public function setAlpha($pValue = 0)
    {
        $this->alpha = $pValue;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5(($this->visible ? 't' : 'f') . $this->blurRadius . $this->distance . $this->direction . $this->alignment . $this->color->getHashCode() . $this->alpha . __CLASS__);
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
