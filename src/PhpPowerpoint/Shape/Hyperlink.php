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

namespace PhpOffice\PhpPowerpoint\Shape;

/**
 * Hyperlink element
 */
class Hyperlink
{
    /**
     * URL to link the shape to
     *
     * @var string
     */
    private $url;

    /**
     * Tooltip to display on the hyperlink
     *
     * @var string
     */
    private $tooltip;

    /**
     * Slide number to link to
     *
     * @var int
     */
    private $slideNumber = null;

    /**
     * Slide relation ID (should not be used by user code!)
     *
     * @var string
     */
    public $relationId = null;

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Shape\Hyperlink
     *
     * @param  string    $pUrl     Url to link the shape to
     * @param  string    $pTooltip Tooltip to display on the hyperlink
     * @throws \Exception
     */
    public function __construct($pUrl = '', $pTooltip = '')
    {
        // Initialise member variables
        $this->setUrl($pUrl);
        $this->setTooltip($pTooltip);
    }

    /**
     * Get URL
     *
     * @return string
     */
    public function getUrl()
    {
        return $this->url;
    }

    /**
     * Set URL
     *
     * @param  string                        $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Hyperlink
     */
    public function setUrl($value = '')
    {
        $this->url = $value;

        return $this;
    }

    /**
     * Get tooltip
     *
     * @return string
     */
    public function getTooltip()
    {
        return $this->tooltip;
    }

    /**
     * Set tooltip
     *
     * @param  string                        $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Hyperlink
     */
    public function setTooltip($value = '')
    {
        $this->tooltip = $value;

        return $this;
    }

    /**
     * Get slide number
     *
     * @return int
     */
    public function getSlideNumber()
    {
        return $this->slideNumber;
    }

    /**
     * Set slide number
     *
     * @param  int                           $value
     * @return \PhpOffice\PhpPowerpoint\Shape\Hyperlink
     */
    public function setSlideNumber($value = 1)
    {
        $this->url         = 'ppaction://hlinksldjump';
        $this->slideNumber = $value;

        return $this;
    }

    /**
     * Is this hyperlink internal? (to another slide)
     *
     * @return boolean
     */
    public function isInternal()
    {
        return strpos($this->url, 'ppaction://') !== false;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->url . $this->tooltip . __CLASS__);
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
