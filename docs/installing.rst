.. _setup:

Installing/configuring
======================

Requirements
------------

Mandatory:

-  PHP 5.3+
-  PHP `Zip <http://php.net/manual/en/book.zip.php>`__ extension
-  PHP `XML
   Parser <http://www.php.net/manual/en/xml.installation.php>`__
   extension

Optional PHP extensions:

-  `XMLWriter <http://php.net/manual/en/book.xmlwriter.php>`__

Installation
------------

There are two ways to install PHPPresentation, i.e. via
`Composer <http://getcomposer.org/>`__ or manually by downloading the
library.

Using Composer
~~~~~~~~~~~~~~

To install via Composer, add the following lines to your
``composer.json``:

.. code-block:: json

    {
        "require": {
           "phpoffice/phppresentation": "dev-master"
        }
    }

Manual install
~~~~~~~~~~~~~~

To install manually, `download PHPOffice\PHPPresentation package from
GitHub <https://github.com/PHPOffice/PHPPresentation/archive/master.zip>`__
and `download PHPOffice\Common package from
GitHub <https://github.com/PHPOffice/Common/archive/master.zip>`__.
Extract the package and put the contents to your machine.

.. code-block:: php

    require_once 'path/to/PhpPresentation/src/PhpPresentation/Autoloader.php';
    \PhpOffice\PhpPresentation\Autoloader::register();

    require_once 'path/to/PhpOffice/Common/src/Common/Autoloader.php';
    \PhpOffice\Common\Autoloader::register();

Using samples
-------------

After installation, you can browse and use the samples that we've
provided, either by command line or using browser. If you can access
your PHPPresentation library folder using browser, point your browser to the
``samples`` folder, e.g. ``http://localhost/PhpPresentation/samples/``.
