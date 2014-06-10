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
 * PHPPowerPoint_DocumentProperties
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class DocumentProperties
{
    /**
     * Creator
     *
     * @var string
     */
    private $_creator;

    /**
     * LastModifiedBy
     *
     * @var string
     */
    private $_lastModifiedBy;

    /**
     * Created
     *
     * @var datetime
     */
    private $_created;

    /**
     * Modified
     *
     * @var datetime
     */
    private $_modified;

    /**
     * Title
     *
     * @var string
     */
    private $_title;

    /**
     * Description
     *
     * @var string
     */
    private $_description;

    /**
     * Subject
     *
     * @var string
     */
    private $_subject;

    /**
     * Keywords
     *
     * @var string
     */
    private $_keywords;

    /**
     * Category
     *
     * @var string
     */
    private $_category;

    /**
     * Company
     *
     * @var string
     */
    private $_company;

    /**
     * Create a new PHPPowerPoint_DocumentProperties
     */
    public function __construct()
    {
        // Initialise values
        $this->_creator        = 'Unknown Creator';
        $this->_lastModifiedBy = $this->_creator;
        $this->_created        = time();
        $this->_modified       = time();
        $this->_title          = "Untitled Presentation";
        $this->_subject        = '';
        $this->_description    = '';
        $this->_keywords       = '';
        $this->_category       = '';
        $this->_company        = 'Microsoft Corporation';
    }

    /**
     * Get Creator
     *
     * @return string
     */
    public function getCreator()
    {
        return $this->_creator;
    }

    /**
     * Set Creator
     *
     * @param  string                           $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setCreator($pValue = '')
    {
        $this->_creator = $pValue;

        return $this;
    }

    /**
     * Get Last Modified By
     *
     * @return string
     */
    public function getLastModifiedBy()
    {
        return $this->_lastModifiedBy;
    }

    /**
     * Set Last Modified By
     *
     * @param  string                           $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setLastModifiedBy($pValue = '')
    {
        $this->_lastModifiedBy = $pValue;

        return $this;
    }

    /**
     * Get Created
     *
     * @return datetime
     */
    public function getCreated()
    {
        return $this->_created;
    }

    /**
     * Set Created
     *
     * @param  datetime                         $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setCreated($pValue = null)
    {
        if (is_null($pValue)) {
            $pValue = time();
        }
        $this->_created = $pValue;

        return $this;
    }

    /**
     * Get Modified
     *
     * @return datetime
     */
    public function getModified()
    {
        return $this->_modified;
    }

    /**
     * Set Modified
     *
     * @param  datetime                         $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setModified($pValue = null)
    {
        if (is_null($pValue)) {
            $pValue = time();
        }
        $this->_modified = $pValue;

        return $this;
    }

    /**
     * Get Title
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->_title;
    }

    /**
     * Set Title
     *
     * @param  string                           $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setTitle($pValue = '')
    {
        $this->_title = $pValue;

        return $this;
    }

    /**
     * Get Description
     *
     * @return string
     */
    public function getDescription()
    {
        return $this->_description;
    }

    /**
     * Set Description
     *
     * @param  string                           $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setDescription($pValue = '')
    {
        $this->_description = $pValue;

        return $this;
    }

    /**
     * Get Subject
     *
     * @return string
     */
    public function getSubject()
    {
        return $this->_subject;
    }

    /**
     * Set Subject
     *
     * @param  string                           $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setSubject($pValue = '')
    {
        $this->_subject = $pValue;

        return $this;
    }

    /**
     * Get Keywords
     *
     * @return string
     */
    public function getKeywords()
    {
        return $this->_keywords;
    }

    /**
     * Set Keywords
     *
     * @param  string                           $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setKeywords($pValue = '')
    {
        $this->_keywords = $pValue;

        return $this;
    }

    /**
     * Get Category
     *
     * @return string
     */
    public function getCategory()
    {
        return $this->_category;
    }

    /**
     * Set Category
     *
     * @param  string                           $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setCategory($pValue = '')
    {
        $this->_category = $pValue;

        return $this;
    }

    /**
     * Get Company
     *
     * @return string
     */
    public function getCompany()
    {
        return $this->_company;
    }

    /**
     * Set Company
     *
     * @param  string                           $pValue
     * @return PHPPowerPoint_DocumentProperties
     */
    public function setCompany($pValue = '')
    {
        $this->_company = $pValue;

        return $this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
}
