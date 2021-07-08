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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation;

class DocumentProperties
{
    /**
     * Creator.
     *
     * @var string
     */
    private $creator = 'Unknown Creator';

    /**
     * LastModifiedBy.
     *
     * @var string
     */
    private $lastModifiedBy = 'Unknown Creator';

    /**
     * Created.
     *
     * @var int
     */
    private $created;

    /**
     * Modified.
     *
     * @var int
     */
    private $modified;

    /**
     * Title.
     *
     * @var string
     */
    private $title = 'Untitled Presentation';

    /**
     * Description.
     *
     * @var string
     */
    private $description = '';

    /**
     * Subject.
     *
     * @var string
     */
    private $subject = '';

    /**
     * Keywords.
     *
     * @var string
     */
    private $keywords = '';

    /**
     * Category.
     *
     * @var string
     */
    private $category = '';

    /**
     * Company.
     *
     * @var string
     */
    private $company = 'Unknown Company';

    /**
     * Generator.
     *
     * @var string
     */
    private $generator = '';

    public function __construct()
    {
        $this->created = time();
        $this->modified = time();
    }

    /**
     * Get Creator.
     *
     * @return string
     */
    public function getCreator(): string
    {
        return $this->creator;
    }

    /**
     * Set Creator.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setCreator(string $pValue = ''): self
    {
        $this->creator = $pValue;

        return $this;
    }

    /**
     * Get Last Modified By.
     *
     * @return string
     */
    public function getLastModifiedBy(): string
    {
        return $this->lastModifiedBy;
    }

    /**
     * Set Last Modified By.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setLastModifiedBy(string $pValue = ''): self
    {
        $this->lastModifiedBy = $pValue;

        return $this;
    }

    /**
     * Get Created.
     *
     * @return int
     */
    public function getCreated(): int
    {
        return $this->created;
    }

    /**
     * Set Created.
     *
     * @param int $pValue
     *
     * @return self
     */
    public function setCreated(int $pValue = null): self
    {
        if (is_null($pValue)) {
            $pValue = time();
        }
        $this->created = $pValue;

        return $this;
    }

    /**
     * Get Modified.
     *
     * @return int
     */
    public function getModified(): int
    {
        return $this->modified;
    }

    /**
     * Set Modified.
     *
     * @param int $pValue
     *
     * @return self
     */
    public function setModified(int $pValue = null): self
    {
        if (is_null($pValue)) {
            $pValue = time();
        }
        $this->modified = $pValue;

        return $this;
    }

    /**
     * Get Title.
     *
     * @return string
     */
    public function getTitle(): string
    {
        return $this->title;
    }

    /**
     * Set Title.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setTitle(string $pValue = ''): self
    {
        $this->title = $pValue;

        return $this;
    }

    /**
     * Get Description.
     *
     * @return string
     */
    public function getDescription(): string
    {
        return $this->description;
    }

    /**
     * Set Description.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setDescription(string $pValue = ''): self
    {
        $this->description = $pValue;

        return $this;
    }

    /**
     * Get Subject.
     *
     * @return string
     */
    public function getSubject(): string
    {
        return $this->subject;
    }

    /**
     * Set Subject.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setSubject(string $pValue = ''): self
    {
        $this->subject = $pValue;

        return $this;
    }

    /**
     * Get Keywords.
     *
     * @return string
     */
    public function getKeywords(): string
    {
        return $this->keywords;
    }

    /**
     * Set Keywords.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setKeywords(string $pValue = ''): self
    {
        $this->keywords = $pValue;

        return $this;
    }

    /**
     * Get Category.
     *
     * @return string
     */
    public function getCategory(): string
    {
        return $this->category;
    }

    /**
     * Set Category.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setCategory(string $pValue = ''): self
    {
        $this->category = $pValue;

        return $this;
    }

    /**
     * Get Company.
     *
     * @return string
     */
    public function getCompany(): string
    {
        return $this->company;
    }

    /**
     * Set Company.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setCompany(string $pValue = ''): self
    {
        $this->company = $pValue;

        return $this;
    }

    /**
     * Get Generator.
     *
     * @return string
     */
    public function getGenerator(): string
    {
        return $this->generator;
    }

    /**
     * Set Generator.
     *
     * @param string $pValue
     *
     * @return self
     */
    public function setGenerator(string $pValue = ''): self
    {
        $this->generator = $pValue;

        return $this;
    }
}
