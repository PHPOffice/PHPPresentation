.. _writersreaders:

Writers
=======


ODPresentation
--------------

The name of the writer is ``ODPresentation``.

.. code-block:: php

    $oWriter = IOFactory::createWriter($oPhpPresentation, 'PowerPoint2007');
    $oWriter->save(__DIR__ . '/sample.pptx');

PowerPoint2007
--------------

The name of the writer is ``PowerPoint2007``.

.. code-block:: php

    $oWriter = IOFactory::createWriter($oPhpPresentation, 'PowerPoint2007');
    $oWriter->save(__DIR__ . '/sample.pptx');

You can change the ZIP Adapter for the writer. By default, the ZIP Adapter is ZipArchiveAdapter.

.. code-block:: php

    use PhpOffice\Common\Adapter\Zip\PclZipAdapter;
    use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;

    $oWriter = IOFactory::createWriter($oPhpPresentation, 'PowerPoint2007');
    $oWriter->setZipAdapter(PclZipAdapter);
    $oWriter->save(__DIR__ . '/sample.pptx');

Serialized
----------

The name of the writer is ``Serialized``.

.. code-block:: php

    $oWriter = IOFactory::createWriter($oPhpPresentation, 'Serialized');
    $oWriter->save(__DIR__ . '/sample.phppt');

You can change the ZIP Adapter for the writer. By default, the ZIP Adapter is ZipArchiveAdapter.

.. code-block:: php

    use PhpOffice\Common\Adapter\Zip\PclZipAdapter;
    use PhpOffice\Common\Adapter\Zip\ZipArchiveAdapter;

    $oWriter = IOFactory::createWriter($oPhpPresentation, 'Serialized');
    $oWriter->setZipAdapter(PclZipAdapter);
    $oWriter->save(__DIR__ . '/sample.phppt');

