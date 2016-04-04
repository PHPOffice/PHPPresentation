.. _shapes:

Shapes
======

Shapes are objects that can be added to a slide. There are five types of shapes that can be used, i.e. [rich text](#rich-text), [line](#line), [chart](#chart), [drawing](#drawing), and [table](#table). Read the corresponding section of this manual for detail information of each shape.

Every shapes have common properties that you can set by using fluent interface.

- ``width`` in pixels
- ``height`` in pixels
- ``offsetX`` in pixels
- ``offsetY`` in pixels
- ``rotation`` in degrees
- ``fill`` see *[Fill](#fill)*
- ``border`` see *[Border](#border)*
- ``shadow`` see *[Shadow](#shadow)*
- ``hyperlink``

Example:

.. code-block:: php

	$richtext = $slide->createRichTextShape()
		->setHeight(300)
		->setWidth(600)
		->setOffsetX(170)
		->setOffsetY(180);

Line
----

To create a line, use `createLineShape` method of slide.
