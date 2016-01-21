.. _faq:

Frequently asked questions
==========================

Is this the same with PHPPowerPoint that I found in CodePlex?
-------------------------------------------------------------

No. This one is much better with tons of new features that you can’t
find in PHPPowerPoint 0.1. The development in CodePlex is halted and
switched to GitHub to allow more participation from the crowd. The more
the merrier, right?

I’ve been running PHPPowerPoint from CodePlex flawlessly, but I can’t use the latest PHPPresentation from GitHub. Why?
--------------------------------------------------------------------------------------------------------------------

PHPPresentation requires PHP 5.3+ since 0.2, while PHPPowerPoint 0.1 from CodePlex
can run with PHP 5.2. There’s a lot of new features that we can get from
PHP 5.3 and it’s been around since 2009! You should upgrade your PHP
version to use PHPPresentation 0.2+.

Why am I getting a class not found error?
-----------------------------------------

If you have followed the instructions for either adding this package to your
``composer.json`` or registering the autoloader, then perhaps you forgot to
include a ``use`` statement for the class(es) you are trying to access.

Here's an example that allows you to refer to the ``MemoryDrawing`` class
without having to specify the full class name in your code:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\MemoryDrawing as MemoryDrawing;

If you *have* followed the installation instructions and you *have* added
the necessary ``use`` statements to your code, then maybe you are still
referencing the ``PHPPowerPoint`` classes using the old PEAR/PSR-0 approach.
The 0.1 approach to naming classes used verbose class names to avoid
namespace collisions with other libraries. For example, the ``MemoryDrawing``
class was actually called ``PHPPowerPoint_Shape_MemoryDrawing``. Version
0.2 of the library renamed the classes, moved to a namespaced approach
and switched to the PSR-0 autoloader. Interestingly, old code that was
still referencing classes using the verbose approach *still worked* (which
was pretty cool!). This is because the PSR-0 autoloader was correctly
translating the verbose class references into the correct file name and
location. However, ``PHPPowerPoint`` now relies exclusively on the PSR-4
autoloader, so old code that may have been referencing the classes with
the verbose class names will need to be updated accordingly.

Why PHPPowerPoint become PHPPresentation ?
------------------------------------------
As `Roman Syroeshko noticed us <https://github.com/PHPOffice/PHPPresentation/issues/25>`__, PowerPoint is a `trademark <http://www.microsoft.com/en-us/legal/IntellectualProperty/Trademarks/EN-US.aspx#332b86b0-befe-4b89-862e-d538e2a653e0>`__.
For avoiding any problems with Microsoft, we decide to change the name to a more logic name, with our panel of readers/writers.