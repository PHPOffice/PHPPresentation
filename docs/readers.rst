.. _writersreaders:

Readers
=======

ODPresentation
--------------

The name of the reader is ``ODPresentation``.

.. code-block:: php

    $oReader = IOFactory::createReader('ODPresentation');
    $oReader->load(__DIR__ . '/sample.odp');

PowerPoint97
------------

The name of the reader is ``PowerPoint97``.

.. code-block:: php

    $oReader = IOFactory::createReader('PowerPoint97');
    $oReader->load(__DIR__ . '/sample.ppt');

PowerPoint2007
--------------

The name of the reader is ``PowerPoint2007``.

.. code-block:: php

    $oReader = IOFactory::createReader('PowerPoint2007');
    $oReader->load(__DIR__ . '/sample.pptx');

Serialized
----------

The name of the reader is ``Serialized``.

.. code-block:: php

    $oReader = IOFactory::createReader('Serialized');
    $oReader->load(__DIR__ . '/sample.phppt');
