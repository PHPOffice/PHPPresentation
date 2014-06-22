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

namespace PhpOffice\PhpPowerpoint;

/**
 * \PhpOffice\PhpPowerpoint\DocumentProperties
 */
class DocumentProperties
{
    /**
     * Creator
     *
     * @var string
     */
    private $creator;

    /**
     * LastModifiedBy
     *
     * @var string
     */
    private $lastModifiedBy;

    /**
     * Created
     *
     * @var int
     */
    private $created;

    /**
     * Modified
     *
     * @var int
     */
    private $modified;

    /**
     * Title
     *
     * @var string
     */
    private $title;

    /**
     * Description
     *
     * @var string
     */
    private $description;

    /**
     * Subject
     *
     * @var string
     */
    private $subject;

    /**
     * Keywords
     *
     * @var string
     */
    private $keywords;

    /**
     * Category
     *
     * @var string
     */
    private $category;

    /**
     * Company
     *
     * @var string
     */
    private $company;

    /**
     * Create a new \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function __construct()
    {
        // Initialise values
        $this->creator        = 'Unknown Creator';
        $this->lastModifiedBy = $this->creator;
        $this->created        = time();
        $this->modified       = time();
        $this->title          = "Untitled Presentation";
        $this->subject        = '';
        $this->description    = '';
        $this->keywords       = '';
        $this->category       = '';
        $this->company        = 'Microsoft Corporation';
    }

    /**
     * Get Creator
     *
     * @return string
     */
    public function getCreator()
    {
        return $this->creator;
    }

    /**
     * Set Creator
     *
     * @param  string                           $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setCreator($pValue = '')
    {
        $this->creator = $pValue;

        return $this;
    }

    /**
     * Get Last Modified By
     *
     * @return string
     */
    public function getLastModifiedBy()
    {
        return $this->lastModifiedBy;
    }

    /**
     * Set Last Modified By
     *
     * @param  string                           $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setLastModifiedBy($pValue = '')
    {
        $this->lastModifiedBy = $pValue;

        return $this;
    }

    /**
     * Get Created
     *
     * @return int
     */
    public function getCreated()
    {
        return $this->created;
    }

    /**
     * Set Created
     *
     * @param int $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setCreated($pValue = null)
    {
        if (is_null($pValue)) {
            $pValue = time();
        }
        $this->created = $pValue;

        return $this;
    }

    /**
     * Get Modified
     *
     * @return int
     */
    public function getModified()
    {
        return $this->modified;
    }

    /**
     * Set Modified
     *
     * @param  int                         $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setModified($pValue = null)
    {
        if (is_null($pValue)) {
            $pValue = time();
        }
        $this->modified = $pValue;

        return $this;
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
     * @param  string                           $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setTitle($pValue = '')
    {
        $this->title = $pValue;

        return $this;
    }

    /**
     * Get Description
     *
     * @return string
     */
    public function getDescription()
    {
        return $this->description;
    }

    /**
     * Set Description
     *
     * @param  string                           $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setDescription($pValue = '')
    {
        $this->description = $pValue;

        return $this;
    }

    /**
     * Get Subject
     *
     * @return string
     */
    public function getSubject()
    {
        return $this->subject;
    }

    /**
     * Set Subject
     *
     * @param  string                           $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setSubject($pValue = '')
    {
        $this->subject = $pValue;

        return $this;
    }

    /**
     * Get Keywords
     *
     * @return string
     */
    public function getKeywords()
    {
        return $this->keywords;
    }

    /**
     * Set Keywords
     *
     * @param  string                           $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setKeywords($pValue = '')
    {
        $this->keywords = $pValue;

        return $this;
    }

    /**
     * Get Category
     *
     * @return string
     */
    public function getCategory()
    {
        return $this->category;
    }

    /**
     * Set Category
     *
     * @param  string                           $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setCategory($pValue = '')
    {
        $this->category = $pValue;

        return $this;
    }

    /**
     * Get Company
     *
     * @return string
     */
    public function getCompany()
    {
        return $this->company;
    }

    /**
     * Set Company
     *
     * @param  string                           $pValue
     * @return \PhpOffice\PhpPowerpoint\DocumentProperties
     */
    public function setCompany($pValue = '')
    {
        $this->company = $pValue;

        return $this;
    }
}
