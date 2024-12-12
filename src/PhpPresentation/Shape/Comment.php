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

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\ComparableInterface;
use PhpOffice\PhpPresentation\Shape\Comment\Author;

/**
 * Comment shape.
 */
class Comment extends AbstractShape implements ComparableInterface
{
    /**
     * @var null|Author
     */
    protected $author;

    /**
     * @var int
     */
    protected $dtComment;

    /**
     * @var string
     */
    protected $text;

    public function __construct()
    {
        parent::__construct();
        $this->setDate(time());
    }

    public function getAuthor(): ?Author
    {
        return $this->author;
    }

    public function setAuthor(Author $author): self
    {
        $this->author = $author;

        return $this;
    }

    /**
     * @return int
     */
    public function getDate()
    {
        return $this->dtComment;
    }

    /**
     * @param int $dtComment timestamp of the comment
     *
     * @return Comment
     */
    public function setDate($dtComment)
    {
        $this->dtComment = (int) $dtComment;

        return $this;
    }

    /**
     * @return string
     */
    public function getText()
    {
        return $this->text;
    }

    /**
     * @param string $text
     *
     * @return Comment
     */
    public function setText($text = '')
    {
        $this->text = $text;

        return $this;
    }

    /**
     * Comment has not height.
     *
     * @return null|int
     */
    public function getHeight()
    {
        return null;
    }

    /**
     * Set Height.
     *
     * @return $this
     */
    public function setHeight(int $pValue = 0)
    {
        return $this;
    }

    /**
     * Comment has not width.
     *
     * @return null|int
     */
    public function getWidth()
    {
        return null;
    }

    /**
     * Set Width.
     *
     * @return self
     */
    public function setWidth(int $pValue = 0)
    {
        return $this;
    }
}
