.. _shapes_richtext:

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
- ``bulletStyle`` see *[Bullet](#bullet)*
- ``lineSpacing`` see *[LineSpacing](#linespacing)*
- ``font`` see *[Font](#font)*


Bullet
------

For a paragraph, you can define the bullet style.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
    use PhpOffice\PhpPresentation\Style\Bullet;

    $oParagraph = new Paragraph();
    $oParagraph->getBulletStyle();

With the bullet style, you can define the char, the font, the color and the type.

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
    use PhpOffice\PhpPresentation\Style\Bullet;
    use PhpOffice\PhpPresentation\Style\Color;

    $oParagraph = new Paragraph();
    $oParagraph->getBulletStyle()->setBulletChar('-');
    $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
    $oParagraph->getBulletStyle()->setBulletColor(new Color(Color::COLOR_RED));


LineSpacing
-----------

For a paragraph, you can define the line spacing.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;

    $oParagraph = new Paragraph();
    $oParagraph->setLineSpacing(200);
    $iLineSpacing = $oParagraph->getLineSpacing();


Run
---

For a run, you can define the language.

Example:

.. code-block:: php

    use PhpOffice\PhpPresentation\Shape\RichText\Run;

    $oRun = new Run();
    $oRun->setLanguage('fr-FR');