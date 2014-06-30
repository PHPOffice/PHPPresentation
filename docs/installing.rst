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

There are two ways to install PHPPowerPoint, i.e. via
`Composer <http://getcomposer.org/>`__ or manually by downloading the
library.

Using Composer
~~~~~~~~~~~~~~

To install via Composer, add the following lines to your
``composer.json``:

.. code-block:: json

    {
        "require": {
           "phpoffice/phppowerpoint": "dev-master"
        }
    }

Manual install
~~~~~~~~~~~~~~

To install manually, `download PHPPowerPoint package from
github <https://github.com/PHPOffice/PHPPowerPoint/archive/master.zip>`__.
Extract the package and put the contents to your machine. To use the
library, include ``src/PhpPowerpoint/Autoloader.php`` in your script and
invoke ``Autoloader::register``.

.. code-block:: php

    require_once '/path/to/src/PhpPowerpoint/Autoloader.php';
    \PhpOffice\PhpPowerpoint\Autoloader::register();

Using samples
-------------

After installation, you can browse and use the samples that we've
provided, either by command line or using browser. If you can access
your PHPPowerPoint library folder using browser, point your browser to the
``samples`` folder, e.g. ``http://localhost/PhpPowerpoint/samples/``.
