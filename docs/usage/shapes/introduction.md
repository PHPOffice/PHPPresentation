# Shapes

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
- ``name``
- ``description`` see *[Alternative text](#alternative-text)*

Example:

``` php
<?php
$richtext = $slide->createRichTextShape()
		->setHeight(300)
		->setWidth(600)
		->setOffsetX(170)
		->setOffsetY(180);
```

## Alternative text

The description of a shape is the alternative text that assistive technologies announce in
place of the shape. Give one to every shape that carries information; leave it empty on
shapes that are purely decorative.

``` php
<?php
$richtext = $slide->createRichTextShape()
		->setName('Budget')
		->setDescription('Budget spent to date: 45% of 1.2M EUR');
```

It is written as the `descr` attribute of `p:cNvPr` in PowerPoint2007 files and as the
`svg:desc` element of the shape in ODPresentation files.

## Line

To create a line, use `createLineShape` method of slide.
