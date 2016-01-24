.. _recipes:

Recipes
=======

How to define the zoom of a presentation ?
------------------------------------------

You must define the zoom of your presentation with the method ``setZoom()``

.. code-block:: php

    // Default
    $zoom = $oPHPPresentation->getZoom();
    // $zoom = 1

    // Without parameter
    $oPHPPresentation->setZoom();
    $zoom = $oPHPPresentation->getZoom();
    // $zoom = true

    // Parameter = false
    $oPHPPresentation->setZoom(2.8);
    $zoom = $oPHPPresentation->getZoom();
    // $zoom = 2.8

How to mark a presentation as final ?
-------------------------------------

You must define your presentation as it with the method ``markAsFinal()``

.. code-block:: php

    // Default
    $state = $oPHPPresentation->isMarkedAsFinal();
    // $state = false

    // Without parameter
    $oPHPPresentation->markAsFinal();
    $state = $oPHPPresentation->isMarkedAsFinal();
    // $state = true

    // Parameter = false
    $oPHPPresentation->markAsFinal(false);
    $state = $oPHPPresentation->isMarkedAsFinal();
    // $state = false

    // Parameter = true
    $oPHPPresentation->markAsFinal(true);
    $state = $oPHPPresentation->isMarkedAsFinal();
    // $state = true
