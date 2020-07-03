.. _general:

General usage
=============

Basic example
-------------

The following is a basic example of the PHPPresentation library. More examples
are provided in the `samples
folder <https://github.com/PHPOffice/PHPPresentation/tree/master/samples/>`__.

.. code-block:: php

    require_once 'src/PhpPresentation/Autoloader.php';
    \PhpOffice\PhpPresentation\Autoloader::register();

    $objPHPPresentation = new PhpPresentation();

    // Create slide
    $currentSlide = $objPHPPresentation->getActiveSlide();

    // Create a shape (drawing)
    $shape = $currentSlide->createDrawingShape();
    $shape->setName('PHPPresentation logo')
          ->setDescription('PHPPresentation logo')
          ->setPath('./resources/phppresentation_logo.gif')
          ->setHeight(36)
          ->setOffsetX(10)
          ->setOffsetY(10);
    $shape->getShadow()->setVisible(true)
                       ->setDirection(45)
                       ->setDistance(10);

    // Create a shape (text)
    $shape = $currentSlide->createRichTextShape()
          ->setHeight(300)
          ->setWidth(600)
          ->setOffsetX(170)
          ->setOffsetY(180);
    $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
    $textRun = $shape->createTextRun('Thank you for using PHPPresentation!');
    $textRun->getFont()->setBold(true)
                       ->setSize(60)
                       ->setColor( new Color( 'FFE06B20' ) );

    $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
    $oWriterPPTX->save(__DIR__ . "/sample.pptx");
    $oWriterODP = IOFactory::createWriter($objPHPPresentation, 'ODPresentation');
    $oWriterODP->save(__DIR__ . "/sample.odp");

Document information
--------------------

You can set the document information such as title, creator, and company
name. Use the following functions :

.. code-block:: php

    $properties = $objPHPPresentation->getProperties();
    $properties->setCreator('My name');
    $properties->setCompany('My factory');
    $properties->setTitle('My title');
    $properties->setDescription('My description');
    $properties->setCategory('My category');
    $properties->setLastModifiedBy('My name');
    $properties->setCreated(mktime(0, 0, 0, 3, 12, 2014));
    $properties->setModified(mktime(0, 0, 0, 3, 14, 2014));
    $properties->setSubject('My subject');
    $properties->setKeywords('my, key, word');


Presentation Properties
-----------------------

You can define some properties which are relative to the presentation, like the zoom or the thumbnail.

Comments
````````

You can define if the presentation display or not the comments with the method ``setCommentVisible``.

.. code-block:: php

    $oPresentation = new PhpPresentation();
    $oProperties = $oPresentation->getPresentationProperties();
    // Get the display for comment
    var_export($oProperties->isCommentVisible());
    // Output : false
    // Enable the display for comment
    $oProperties->setCommentVisible(true);
    // Get the display for comment
    var_export($oProperties->isCommentVisible());
    // Output : true

Last View
`````````

You can define the last view of the presentation with the method ``setLastView``.

.. code-block:: php

    $oPresentation = new PhpPresentation();
    $oProperties = $oPresentation->getPresentationProperties();
    // Get the last view of the presentation
    echo $oProperties->getZoom();
    // Output : PresentationProperties::VIEW_SLIDE
    // Set the last view of the presentation
    $oProperties->setLastView(PresentationProperties::VIEW_NOTES);
    // Get the last view of the presentation
    echo $oProperties->getZoom();
    // Output : PresentationProperties::VIEW_NOTES

Thumbnail
`````````

You can define the thumbnail of the presentation with the method ``setThumbnailPath``.

.. code-block:: php

    $oPresentation = new PhpPresentation();
    $oProperties = $oPresentation->getPresentationProperties();
    // Set path of the thumbnail
    $oProperties->setThumbnailPath(__DIR__.'\resources\phppowerpoint_logo.gif');
    // Get path of the thumbnail
    echo $oProperties->getThumbnailPath();

Zoom
````

You can define the zoom of the presentation with the method ``setZoom``.

.. code-block:: php

    $oPresentation = new PhpPresentation();
    $oProperties = $oPresentation->getPresentationProperties();
    // Get zoom of the presentation
    echo $oProperties->getZoom();
    // Output : 1
    // Set zoom of the presentation (3 = 300%)
    $oProperties->setZoom(3);
    // Get zoom of the presentation
    echo $oProperties->getZoom();
    // Output : 3
