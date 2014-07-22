.. _general:

General usage
=============

Basic example
-------------

The following is a basic example of the PHPPowerPoint library. More examples
are provided in the `samples
folder <https://github.com/PHPOffice/PHPPowerPoint/tree/master/samples/>`__.

.. code-block:: php

    require_once 'src/PhpPowerpoint/Autoloader.php';
    \PhpOffice\PhpPowerpoint\Autoloader::register();

    $objPHPPowerPoint = new PhpPowerpoint();

    // Create slide
    $currentSlide = $objPHPPowerPoint->getActiveSlide();

    // Create a shape (drawing)
    $shape = $currentSlide->createDrawingShape();
    $shape->setName('PHPPowerPoint logo')
          ->setDescription('PHPPowerPoint logo')
          ->setPath('./resources/phppowerpoint_logo.gif')
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
    $textRun = $shape->createTextRun('Thank you for using PHPPowerPoint!');
    $textRun->getFont()->setBold(true)
                       ->setSize(60)
                       ->setColor( new Color( 'FFE06B20' ) );

    $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
    $oWriterPPTX->save(__DIR__ . "/sample.pptx");
    $oWriterODP = IOFactory::createWriter($objPHPPowerPoint, 'ODPresentation');
    $oWriterODP->save(__DIR__ . "/sample.odp");

Document information
--------------------

You can set the document information such as title, creator, and company
name. Use the following functions:

.. code-block:: php

    $properties = $phpPowerpoint->getProperties();
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

