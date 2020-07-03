.. _shapes_comment:

Comments
========

To create a comment, create an object `Comment`.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Comment;

    $oComment = new Comment();
    $oSlide->addShape($oComment);

You can define text and date with setters.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Comment;

    $oComment = new Comment();
    $oComment->setText('Text of the Comment');
    $oComment->setDate(time());
    $oSlide->addShape($oComment);


Author
------

For a comment, you can define the author.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Comment;
    use PhpOffice\PhpPresentation\Shape\Comment\Author;

    $oAuthor = new Author();
    $oComment = new Comment();
    $oComment->setAuthor($oAuthor);
    $oSlide->addShape($oComment);

You can define name and initials with setters.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Comment;
    use PhpOffice\PhpPresentation\Shape\Comment\Author;

    $oAuthor = new Author();
    $oAuthor->setName('Name of the author');
    $oAuthor->setInitals('Nota');
    $oComment = new Comment();
    $oComment->setAuthor($oAuthor);
    $oSlide->addShape($oComment);