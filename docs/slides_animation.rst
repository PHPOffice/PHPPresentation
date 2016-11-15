.. _slides_animation:

Animation
=========

You can create multiples animations in a slide.

.. code-block:: php

    use PhpOffice\PhpPresentation\Slide\Animation;

    $oAnimation1 = new Animation();
    $oAnimation1->addShape($oDrawing);
    $oSlide->addAnimation($oAnimation1);

    $oAnimation2 = new Animation();
    $oAnimation2->addShape($oRichText);
    $oSlide->addAnimation($oAnimation2);

