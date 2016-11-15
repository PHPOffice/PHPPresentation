.. _slides:

Slides
======

Slides are pages in a presentation. Slides are stored as a zero based array in ``PHPPresentation`` object. Use ``createSlide`` to create a new slide and retrieve the slide for other operation such as creating shapes for that slide.

Name
----

By default, a slide has not a name.
You can define it with the method ``setName``.

.. code-block:: php

	$oSlide = $oPHPPresentation->createSlide();
	$oSlide->setName('Title of the slide');

Visibility
----------

By default, a slide is visible.
You can define it with the method ``setIsVisible``.

.. code-block:: php

	$oSlide = $oPHPPresentation->createSlide();
	$oSlide->setIsVisible(false);
	var_dump($oSlide->isVisible());