.. _shapes_chart:

RichText
========

Rich text shapes contain paragraphs of texts. To create a rich text shape, use ``createRichTextShape`` method of slide.

Below are the properties that you can set for a rich text shape.

- ``wrap``
- ``autoFit``
- ``fontScale`` : font scale (in percentage) when autoFit = RichText::AUTOFIT_NORMAL
- ``lnSpcReduction`` : line spacing reduction (in percentage) when autoFit = RichText::AUTOFIT_NORMAL
- ``horizontalOverflow``
- ``verticalOverflow``
- ``upright``
- ``vertical``
- ``columns``
- ``bottomInset`` in pixels
- ``leftInset`` in pixels
- ``rightInset`` in pixels
- ``topInset`` in pixels
- ``autoShrinkHorizontal`` (boolean)
- ``autoShrinkVertical`` (boolean)

Properties that can be set for each paragraphs are as follow.

- ``alignment`` see *[Alignment](#alignment)*
- ``font`` see *[Font](#font)*
- ``bulletStyle`` see *[Bullet](#bullet)*


Run
---

For a run, you can define the language.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\RichText\Run;

    $oRun = new Run();
    $oComment->setLanguage('fr-FR');