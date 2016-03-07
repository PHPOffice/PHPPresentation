<?php

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
     * @return Author
     */
    public function setName($name)
    {
        $this->name = $name;
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string Hash code
     */
    public function getHashCode()
    {
        return md5($this->getInitials() . $this->getName() . __CLASS__);
    }
}
