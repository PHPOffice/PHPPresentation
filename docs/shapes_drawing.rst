.. _shapes_drawing:

Drawing
=======

To create a drawing, you have four sources : File, GD, Base64 and ZipFile.

File
----

To create a drawing, use `createDrawingShape` method of slide.

.. code-block:: php

	$oShape = $oSlide->createDrawingShape();
	$oShape->setName('Unique name')
		->setDescription('Description of the drawing')
		->setPath('/path/to/drawing.filename');

It's an alias for :

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Drawing\File;

    $oShape = new File();
	$oShape->setName('Unique name')
		->setDescription('Description of the drawing')
		->setPath('/path/to/drawing.filename');
    $oSlide->addShape($oShape);

GD
--

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Drawing\Gd;

    $gdImage = @imagecreatetruecolor($width, $height);

    $oShape = new Gd();
    $oShape->setName('Sample image')
        ->setDescription('Sample image')
        ->setImageResource($gdImage)
        ->setRenderingFunction(Drawing\Gd::RENDERING_JPEG)
        ->setMimeType(Drawing\Gd::MIMETYPE_DEFAULT);
    $oSlide->addShape($oShape);


Base64
------

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Drawing\Base64;

    $oShape = new Base64();
    $oShape->setName('Sample image')
        ->setDescription('Sample image')
        ->setImageResource($gdImage)
        ->setData('data:image/jpeg;base64,..........');
    $oSlide->addShape($oShape);

ZipFile
-------

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\Drawing\ZipFile;

    $oShape = new ZipFile();
    $oShape->setName('Sample image')
        ->setDescription('Sample image')
        ->setPath('zip://myzipfile.zip#path/in/zip/img.ext')
    $oSlide->addShape($oShape);
