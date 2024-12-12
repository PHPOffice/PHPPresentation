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

namespace PhpOffice\PhpPresentation\Shape\Comment;

class Author
{
    /**
     * @var int
     */
    protected $idxAuthor;

    /**
     * @var string
     */
    protected $initials;

    /**
     * @var string
     */
    protected $name;

    /**
     * @return int
     */
    public function getIndex()
    {
        return $this->idxAuthor;
    }

    /**
     * @param int $idxAuthor
     *
     * @return Author
     */
    public function setIndex($idxAuthor)
    {
        $this->idxAuthor = (int) $idxAuthor;

        return $this;
    }

    /**
     * @return mixed
     */
    public function getInitials()
    {
        return $this->initials;
    }

    /**
     * @param mixed $initials
     *
     * @return Author
     */
    public function setInitials($initials)
    {
        $this->initials = $initials;

        return $this;
    }

    /**
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * @param string $name
     *
     * @return Author
     */
    public function setName($name)
    {
        $this->name = $name;

        return $this;
    }

    /**
     * Get hash code.
     *
     * @return string Hash code
     */
    public function getHashCode(): string
    {
        return md5($this->getInitials() . $this->getName() . __CLASS__);
    }
}
