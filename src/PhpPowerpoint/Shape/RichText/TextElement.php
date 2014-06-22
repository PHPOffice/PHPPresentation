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

namespace PhpOffice\PhpPowerpoint\Shape\RichText;

use PhpOffice\PhpPowerpoint\Shape\Hyperlink;

/**
 * Rich text text element
 */
class TextElement implements TextElementInterface
{
    /**
     * Text
     *
     * @var string
     */
    private $text;

    /**
     * Hyperlink
     *
     * @var \PhpOffice\PhpPowerpoint\Shape\Hyperlink
     */
    protected $hyperlink;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Shape\RichText\TextElement instance
     *
     * @param string $pText Text
     */
    public function __construct($pText = '')
    {
        // Initialise variables
        $this->text = $pText;
    }

    /**
     * Get text
     *
     * @return string Text
     */
    public function getText()
    {
        return $this->text;
    }

    /**
     * Set text
     *
     * @param                                            $pText string   Text
     * @return \PhpOffice\PhpPowerpoint\Shape\RichText\TextElementInterface
     */
    public function setText($pText = '')
    {
        $this->text = $pText;

        return $this;
    }

    /**
     * Get font
     *
     * @return \PhpOffice\PhpPowerpoint\Style\Font
     */
    public function getFont()
    {
        return null;
    }

    /**
     * Has Hyperlink?
     *
     * @return boolean
     */
    public function hasHyperlink()
    {
        return !is_null($this->hyperlink);
    }

    /**
     * Get Hyperlink
     *
     * @return \PhpOffice\PhpPowerpoint\Shape\Hyperlink
     */
    public function getHyperlink()
    {
        if (is_null($this->hyperlink)) {
            $this->hyperlink = new Hyperlink();
        }

        return $this->hyperlink;
    }

    /**
     * Set Hyperlink
     *
     * @param  \PhpOffice\PhpPowerpoint\Shape\Hyperlink $pHyperlink
     * @throws \Exception
     * @return \PhpOffice\PhpPowerpoint\AbstractShape
     */
    public function setHyperlink(Hyperlink $pHyperlink = null)
    {
        $this->hyperlink = $pHyperlink;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->text . (is_null($this->hyperlink) ? '' : $this->hyperlink->getHashCode()) . __CLASS__);
    }
}
