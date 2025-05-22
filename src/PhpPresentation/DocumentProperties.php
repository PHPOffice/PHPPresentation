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

namespace PhpOffice\PhpPresentation;

class DocumentProperties
{
    public const PROPERTY_TYPE_BOOLEAN = 'b';
    public const PROPERTY_TYPE_INTEGER = 'i';
    public const PROPERTY_TYPE_FLOAT = 'f';
    public const PROPERTY_TYPE_DATE = 'd';
    public const PROPERTY_TYPE_STRING = 's';
    public const PROPERTY_TYPE_UNKNOWN = 'u';

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
     * Revision.
     *
     * @var string
     */
    private $revision = '';

    /**
     * Status.
     *
     * @var string
     */
    private $status = '';

    /**
     * Custom Properties.
     *
     * @var array<string, array<string, mixed>>
     */
    private $customProperties = [];

    /**
     * Generator.
     *
     * @var string
     */
    private $generator = '';

    /**
     * Create a new \PhpOffice\PhpPresentation\DocumentProperties.
     */
    public function __construct()
    {
        $this->created = time();
        $this->modified = time();
    }

    /**
     * Get Creator.
     */
    public function getCreator(): string
    {
        return $this->creator;
    }

    /**
     * Set Creator.
     */
    public function setCreator(string $pValue = ''): self
    {
        $this->creator = $pValue;

        return $this;
    }

    /**
     * Get Last Modified By.
     */
    public function getLastModifiedBy(): string
    {
        return $this->lastModifiedBy;
    }

    /**
     * Set Last Modified By.
     */
    public function setLastModifiedBy(string $pValue = ''): self
    {
        $this->lastModifiedBy = $pValue;

        return $this;
    }

    /**
     * Get Created.
     */
    public function getCreated(): int
    {
        return $this->created;
    }

    /**
     * Set Created.
     */
    public function setCreated(?int $pValue = null): self
    {
        if (null === $pValue) {
            $pValue = time();
        }
        $this->created = $pValue;

        return $this;
    }

    /**
     * Get Modified.
     */
    public function getModified(): int
    {
        return $this->modified;
    }

    /**
     * Set Modified.
     */
    public function setModified(?int $pValue = null): self
    {
        if (null === $pValue) {
            $pValue = time();
        }
        $this->modified = $pValue;

        return $this;
    }

    /**
     * Get Title.
     */
    public function getTitle(): string
    {
        return $this->title;
    }

    /**
     * Set Title.
     */
    public function setTitle(string $pValue = ''): self
    {
        $this->title = $pValue;

        return $this;
    }

    /**
     * Get Description.
     */
    public function getDescription(): string
    {
        return $this->description;
    }

    /**
     * Set Description.
     */
    public function setDescription(string $pValue = ''): self
    {
        $this->description = $pValue;

        return $this;
    }

    /**
     * Get Subject.
     */
    public function getSubject(): string
    {
        return $this->subject;
    }

    /**
     * Set Subject.
     */
    public function setSubject(string $pValue = ''): self
    {
        $this->subject = $pValue;

        return $this;
    }

    /**
     * Get Keywords.
     */
    public function getKeywords(): string
    {
        return $this->keywords;
    }

    /**
     * Set Keywords.
     */
    public function setKeywords(string $pValue = ''): self
    {
        $this->keywords = $pValue;

        return $this;
    }

    /**
     * Get Category.
     */
    public function getCategory(): string
    {
        return $this->category;
    }

    /**
     * Set Category.
     */
    public function setCategory(string $pValue = ''): self
    {
        $this->category = $pValue;

        return $this;
    }

    /**
     * Get Company.
     */
    public function getCompany(): string
    {
        return $this->company;
    }

    /**
     * Set Company.
     */
    public function setCompany(string $pValue = ''): self
    {
        $this->company = $pValue;

        return $this;
    }

    /**
     * Get a List of Custom Property Names.
     *
     * @return array<int, string>
     */
    public function getCustomProperties(): array
    {
        return array_keys($this->customProperties);
    }

    /**
     * Check if a Custom Property is defined.
     */
    public function isCustomPropertySet(string $propertyName): bool
    {
        return isset($this->customProperties[$propertyName]);
    }

    /**
     * Get a Custom Property Value.
     *
     * @return null|mixed
     */
    public function getCustomPropertyValue(string $propertyName)
    {
        if ($this->isCustomPropertySet($propertyName)) {
            return $this->customProperties[$propertyName]['value'];
        }

        return null;
    }

    /**
     * Get a Custom Property Type.
     */
    public function getCustomPropertyType(string $propertyName): ?string
    {
        if ($this->isCustomPropertySet($propertyName)) {
            return $this->customProperties[$propertyName]['type'];
        }

        return null;
    }

    /**
     * Set a Custom Property.
     *
     * @param mixed $propertyValue
     * @param null|string $propertyType
     *                                  'i' : Integer
     *                                  'f' : Floating Point
     *                                  's' : String
     *                                  'd' : Date/Time
     *                                  'b' : Boolean
     */
    public function setCustomProperty(string $propertyName, $propertyValue = '', ?string $propertyType = null): self
    {
        if (!in_array($propertyType, [
            self::PROPERTY_TYPE_INTEGER,
            self::PROPERTY_TYPE_FLOAT,
            self::PROPERTY_TYPE_STRING,
            self::PROPERTY_TYPE_DATE,
            self::PROPERTY_TYPE_BOOLEAN,
        ])) {
            if (is_float($propertyValue)) {
                $propertyType = self::PROPERTY_TYPE_FLOAT;
            } elseif (is_int($propertyValue)) {
                $propertyType = self::PROPERTY_TYPE_INTEGER;
            } elseif (is_bool($propertyValue)) {
                $propertyType = self::PROPERTY_TYPE_BOOLEAN;
            } else {
                $propertyType = self::PROPERTY_TYPE_STRING;
            }
        }
        $this->customProperties[$propertyName] = [
            'value' => $propertyValue,
            'type' => $propertyType,
        ];

        return $this;
    }

    /**
     * Get Revision.
     */
    public function getRevision(): string
    {
        return $this->revision;
    }

    /**
     * Set Revision.
     */
    public function setRevision(string $pValue = ''): self
    {
        $this->revision = $pValue;

        return $this;
    }

    /**
     * Get Status.
     */
    public function getStatus(): string
    {
        return $this->status;
    }

    /**
     * Set Status.
     */
    public function setStatus(string $pValue = ''): self
    {
        $this->status = $pValue;

        return $this;
    }

    /**
     * Get Generator.
     */
    public function getGenerator(): string
    {
        return $this->generator;
    }

    /**
     * Set Generator.
     */
    public function setGenerator(string $pValue = ''): self
    {
        $this->generator = $pValue;

        return $this;
    }
}
