# Comments

To create a comment, create an object `Comment`.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Comment;

$comment = new Comment();
$slide->addShape($comment);
```

You can define text and date with setters.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Comment;

$comment = new Comment();
$comment->setText('Text of the Comment');
$comment->setDate(time());
$slide->addShape($comment);
```

## Author

For a comment, you can define the author.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Comment\Author;

$author = new Author();
$comment = new Comment();
$comment->setAuthor($author);
$slide->addShape($comment);
```

You can define name and initials with setters.

Example:

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Comment\Author;

$author = new Author();
$author->setName('Name of the author');
$author->setInitals('Nota');
$comment = new Comment();
$comment->setAuthor($author);
$slide->addShape($comment);
```