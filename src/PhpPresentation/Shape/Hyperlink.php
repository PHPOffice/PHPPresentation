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

namespace PhpOffice\PhpPresentation\Shape;

/**
 * Hyperlink element.
 */
class Hyperlink
{
    /**
     * URL to link the shape to.
     *
     * @var string
     */
    private $url;

    /**
     * Tooltip to display on the hyperlink.
     *
     * @var string
     */
    private $tooltip;

    /**
     * Slide number to link to.
     *
     * @var int
     */
    private $slideNumber;

    /**
     * Slide relation ID (should not be used by user code!).
     *
     * @var string
     */
    public $relationId;

    /**
     * Hash index.
     *
     * @var int
     */
    private $hashIndex;

    /**
     * If true, uses the text color, instead of theme color.
     *
     * @var bool
     */
    private $isTextColorUsed = false;

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Hyperlink.
     *
     * @param string $pUrl Url to link the shape to
     * @param string $pTooltip Tooltip to display on the hyperlink
     */
    public function __construct(string $pUrl = '', string $pTooltip = '')
    {
        $this->setUrl($pUrl);
        $this->setTooltip($pTooltip);
    }

    /**
     * Get URL.
     */
    public function getUrl(): string
    {
        return $this->url;
    }

    /**
     * Set URL.
     */
    public function setUrl(string $value = ''): self
    {
        $this->url = $value;

        return $this;
    }

    /**
     * Get tooltip.
     */
    public function getTooltip(): string
    {
        return $this->tooltip;
    }

    /**
     * Set tooltip.
     */
    public function setTooltip(string $value = ''): self
    {
        $this->tooltip = $value;

        return $this;
    }

    /**
     * Get slide number.
     */
    public function getSlideNumber(): int
    {
        return $this->slideNumber;
    }

    /**
     * Set slide number.
     */
    public function setSlideNumber(int $value = 1): self
    {
        $this->url = 'ppaction://hlinksldjump';
        $this->slideNumber = $value;

        return $this;
    }

    /**
     * Is this hyperlink internal? (to another slide).
     */
    public function isInternal(): bool
    {
        return false !== strpos($this->url, 'ppaction://');
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5($this->url . $this->tooltip . __CLASS__);
    }

    /**
     * Get hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @return null|int Hash index
     */
    public function getHashIndex(): ?int
    {
        return $this->hashIndex;
    }

    /**
     * Set hash index.
     *
     * Note that this index may vary during script execution! Only reliable moment is
     * while doing a write of a workbook and when changes are not allowed.
     *
     * @param int $value Hash index
     *
     * @return $this
     */
    public function setHashIndex(int $value)
    {
        $this->hashIndex = $value;

        return $this;
    }

    /**
     * Get whether or not to use text color for a hyperlink, instead of theme color.
     *
     * @see https://docs.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/014fbc20-3705-4812-b8cd-93f5af05b504
     *
     * @return bool whether or not to use text color for a hyperlink, instead of theme color
     */
    public function isTextColorUsed(): bool
    {
        return $this->isTextColorUsed;
    }

    /**
     * Set whether or not to use text color for a hyperlink, instead of theme color.
     *
     * @see https://docs.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/014fbc20-3705-4812-b8cd-93f5af05b504
     */
    public function setIsTextColorUsed(bool $isTextColorUsed): self
    {
        $this->isTextColorUsed = $isTextColorUsed;

        return $this;
    }
}
