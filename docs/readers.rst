.. _writersreaders:

Readers
=======

ODPresentation
--------------

The name of the reader is ``ODPresentation``.

.. code-block:: php

    $oWriter = IOFactory::createReader('ODPresentation');
    $oWriter->load(__DIR__ . '/sample.odp');

PowerPoint97
------------

The name of the reader is ``PowerPoint97``.

.. code-block:: php

    $oWriter = IOFactory::createReader('PowerPoint97');
    $oWriter->load(__DIR__ . '/sample.ppt');

PowerPoint2007
--------------

The name of the reader is ``PowerPoint2007``.

.. code-block:: php

    $oWriter = IOFactory::createReader('PowerPoint2007');
    $oWriter->load(__DIR__ . '/sample.pptx');

Serialized
----------

The name of the reader is ``Serialized``.

.. code-block:: php

    $oWriter = IOFactory::createReader('Serialized');
    $oWriter->load(__DIR__ . '/sample.phppt');
