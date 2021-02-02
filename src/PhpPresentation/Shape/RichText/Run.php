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

namespace PhpOffice\PhpPresentation\Shape\RichText;

use PhpOffice\PhpPresentation\Style\Font;

/**
 * Rich text run.
 */
class Run extends TextElement implements TextElementInterface
{
    /**
     * Font.
     *
     * @var \PhpOffice\PhpPresentation\Style\Font
     */
    private $font;

    /**
     * List of effect apply to paragraph
     * @var array \PhpOffice\PhpPresentation\Style\Effect[]
     */
    protected ?array $effectCollection = null;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\RichText\Run instance
     *
     * @param string $pText Text
     */
    public function __construct($pText = '')
    {
        // Initialise variables
        $this->setText($pText);
        $this->font = new Font();
        $this->effectCollection = null;
    }

    /**
     * Get font.
     */
    public function getFont(): Font
    {
        return $this->font;
    }

    /**
     * Set font.
     *
     * @param null|Font $pFont Font
     *
     * @return \PhpOffice\PhpPresentation\Shape\RichText\TextElementInterface
     */
    public function setFont(?Font $pFont = null)
    {
        $this->font = $pFont;

        return $this;
    }

    /**
     * Add an effect to the shpae
     * 
     * @param \PhpOffice\PhpPresentation\Style\Effect $effect
     * @return $this
     */
    public function addEffect(Shape\Effect $effect)
    {
      if (!isset($this->effectCollection)) {
        $this->effectCollection = array();
      }
      $this->effectCollection[] = $effect;
      return $this;
    }
    
    /**
     * Get the effect collection
     * 
     * @return array \PhpOffice\PhpPresentation\Style\Effect[]
     */
    public function getEffectCollection():?array
    {
      return $this->effectCollection;
    }
    
    /**
     * Set the effect collection
     * 
     * @param array \PhpOffice\PhpPresentation\Style\Effect $effectCollection
     * @return $this
     */
    public function setEffectCollection(array $effectCollection)
    {
      if (   isset($effectCollection)
          && is_array($effectCollection)) {
        $this->effectCollection = $effectCollection;
      }
      return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5($this->getText() . $this->font->getHashCode() . __CLASS__);
    }
}
