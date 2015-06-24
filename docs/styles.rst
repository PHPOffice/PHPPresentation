.. _styles:

Styles
======

Fill
----

Use this style to define fill of a shape as example below.

.. code-block:: php

	$shape->getFill()
		->setFillType(Fill::FILL_GRADIENT_LINEAR)
		->setRotation(270)
		->setStartColor(new Color('FFCCCCCC'))
		->setEndColor(new Color('FFFFFFFF'));

Properties:

- ``fillType``
- ``rotation``
- ``startColor``
- ``endColor``

Border
------

Use this style to define border of a shape as example below.

.. code-block:: php

	$shape->getBorder()
		->setLineStyle(Border::LINE_SINGLE)
		->setLineWidth(4)
		->getColor()->setARGB('FFC00000');

Properties:

- ``lineWidth``
- ``lineStyle``
- ``dashStyle``
- ``color``

Shadow
------

Use this style to define shadow of a shape as example below.

.. code-block:: php

	$shape->getShadow()
		->setVisible(true)
		->setDirection(45)
		->setDistance(10);

Properties:

- ``visible``
- ``blurRadius``
- ``distance``
- ``direction``
- ``alignment``
- ``color``
- ``alpha``

Alignment
---------

- ``horizontal``
- ``vertical``
- ``level``
- ``indent``
- ``marginLeft``
- ``marginRight``

Font
----

- ``name``
- ``bold``
- ``italic``
- ``superScript``
- ``subScript``
- ``underline``
- ``strikethrough``
- ``color``

Bullet
------

- ``bulletType``
- ``bulletFont``
- ``bulletChar``
- ``bulletNumericStyle``
- ``bulletNumericStartAt``

Color
-----

Colors can be applied to different objects, e.g. font or border.

.. code-block:: php

	$textRun = $shape->createTextRun('Text');
	$textRun->getFont()->setColor(new Color('C00000'));
