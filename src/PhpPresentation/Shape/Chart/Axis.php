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

namespace PhpOffice\PhpPresentation\Shape\Chart;

use PhpOffice\PhpPresentation\ComparableInterface;

/**
 * \PhpOffice\PhpPresentation\Shape\Chart\Axis
 */
class Axis implements ComparableInterface
{
    /**
     * Title
     *
     * @var string
     */
    private $title = 'Axis Title';

    /**
     * Format code
     *
     * @var string
     */
    private $formatCode = '';

    /**
     * Create a new \PhpOffice\PhpPresentation\Shape\Chart\Axis instance
     *
     * @param string $title Title
     */
    public function __construct($title = 'Axis Title')
    {
        $this->title = $title;
    }

    /**
     * Get Title
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->title;
    }

    /**
     * Set Title
     *
     * @param  string                         $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    public function setTitle($value = 'Axis Title')
    {
        $this->title = $value;

        return $this;
    }

    /**
     * Get Format Code
     *
     * @return string
     */
    public function getFormatCode()
    {
        return $this->formatCode;
    }

    /**
     * Set Format Code
     *
     * @param  string                         $value
     * @return \PhpOffice\PhpPresentation\Shape\Chart\Axis
     */
    public function setFormatCode($value = '')
    {
        $this->formatCode = $value;

        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->title . $this->formatCode . __CLASS__);
    }

    /**
     * Hash index
     *
     * @var string
     */
    private $hashIndex;

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
        return $this;
    }
}
