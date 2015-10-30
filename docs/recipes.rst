.. _recipes:

Recipes
=======

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
