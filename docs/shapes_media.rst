.. _shapes_table:

Media
=====

To create a video, create an object `Media`.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Media;

    $oMedia = new Media();
    $oMedia->setPath('file.mp4');
    // $oMedia->setPath('file.ogv');
    $oSlide->addShape($oMedia);

You can define text and date with setters.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Media;

    $oMedia = new Media();
    $oMedia->setName('Name of the Media');
    $oSlide->addShape($oMedia);


Quirks
------

For PowerPoint2007 Writer, the prefered file format is MP4.
For ODPresentation Writer, the prefered file format is OGV.