.. _shapes_chart:

Charts
======

To create a chart, use `createChartShape` method of Slide.

Example:

.. code-block:: php

    $chartShape = $slide->createChartShape();

Parts
-----

Axis
^^^^

You can define gridlines (minor and major) for each axis (X & Y).
For each gridline, you can custom the width (in points), the fill type and the fill color.

.. code-block:: php

    use \PhpOffice\PhpPresentation\Shape\Chart\Gridlines;

    $oLine = new Line();

    $oGridLines = new Gridlines();
    $oGridLines->getOutline()->setWidth(10);
    $oGridLines->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));

    $oShape = $oSlide->createChartShape();
    $oShape->getPlotArea()->setType($oLine);
    $oShape->getPlotArea()->getAxisX()->setMajorGridlines($oGridLines);

For Axis, you can define the min & max bounds with `setMinBounds` & `setMaxBounds` methods.
For resetting them, you pass null as parameter to these methods.

.. code-block:: php

    use \PhpOffice\PhpPresentation\Shape\Chart\Gridlines;

    $oLine = new Line();

    $oShape = $oSlide->createChartShape();
    $oShape->getPlotArea()->setType($oLine);
    $oShape->getPlotArea()->getAxisX()->setMinBounds(0);
    $oShape->getPlotArea()->getAxisX()->setMaxBounds(200);

You can define outline for each axis (X & Y).

.. code-block:: php

    $oLine = new Line();

    $oShape = $oSlide->createChartShape();
    $oShape->getPlotArea()->setType($oLine);
    $oShape->getPlotArea()->getAxisX()->getOutline()->setWidth(10);
    $oShape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));

For Axis Y, you can define tick mark with `setMinorTickMark` & `setMajorTickMark` methods.
For resetting them, you pass Axis::TICK_MARK_NONE as parameter to these methods.

.. code-block:: php

    use \PhpOffice\PhpPresentation\Shape\Chart\Axis;

    $oLine = new Line();

    $oShape = $oSlide->createChartShape();
    $oShape->getPlotArea()->setType($oLine);
    $oShape->getPlotArea()->getAxisY()->setMinorTickMark(Axis::TICK_MARK_NONE);
    $oShape->getPlotArea()->getAxisY()->setMajorTickMark(Axis::TICK_MARK_INSIDE);

For Axis Y, you can define unit with `setMinorUnit` & `setMajorUnit` methods.
For resetting them, you pass null as parameter to these methods.

.. code-block:: php

    use \PhpOffice\PhpPresentation\Shape\Chart\Axis;

    $oLine = new Line();

    $oShape = $oSlide->createChartShape();
    $oShape->getPlotArea()->setType($oLine);
    $oShape->getPlotArea()->getAxisY()->setMinorUnit(null);
    $oShape->getPlotArea()->getAxisY()->setMajorUnit(0.05);

You can define visibility for each axis (X & Y).

.. code-block:: php

    $oLine = new Line();

    $oShape = $oSlide->createChartShape();
    $oShape->getPlotArea()->setType($oLine);
    $oShape->getPlotArea()->getAxisX()->setIsVisible(false);

Title
^^^^^

By default, the title of a chart is displayed. 
For hiding it, you define its visibility to false.

.. code-block:: php

    $oLine = new Line();
    $oShape = $slide->createChartShape();
    $oShape->getPlotArea()->setType($oLine);
    // Hide the title
    $oShape->getTitle()->setVisible(false);

Series
^^^^^^

You can custom the font of a serie.

.. code-block:: php
    $oSeries = new Series('Downloads', $seriesData);
    // Define the size
    $oSeries->getFont()->setSize(25);

You can custom the marker of a serie, for Line & Scatter charts.

.. code-block:: php
    use \PhpOffice\PhpPresentation\Shape\Chart\Marker;

    $oSeries = new Series('Downloads', $seriesData);
    $oMarker = $oSeries->getMarker();
    $oMarker->setSymbol(Marker::SYMBOL_DASH)->setSize(10);

You can custom the line of a serie, for Line & Scatter charts.

.. code-block:: php
    use \PhpOffice\PhpPresentation\Style\Outline;

    $oOutline = new Outline();
    // Define the color
    $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
    $oOutline->getFill()->setStartColor(new Color(Color::COLOR_YELLOW));
    // Define the width (in points)
    $oOutline->setWidth(2);

    $oSeries = new Series('Downloads', $seriesData);
    $oSeries->setOutline($oOutline);

You can define the position of the data label.
Each position is described in `MSDN <https://msdn.microsoft.com/en-us/library/mt459417(v=office.12).aspx>`_

.. code-block:: php

    $oSeries = new Series('Downloads', $seriesData);
    $oSeries->setLabelPosition(Series::LABEL_INSIDEEND);

You can define if some informations are displayed.

.. code-block:: php

    $oSeries = new Series('Downloads', $seriesData);
    $oSeries->setSeparator(';');
    $oSeries->setShowCategoryName(true);
    $oSeries->setShowLeaderLines(true);
    $oSeries->setShowLegendKey(true);
    $oSeries->setShowPercentage(true);
    $oSeries->setShowSeriesName(true);
    $oSeries->setShowValue(true);

View3D
^^^^^^

For enabling the autoscale for a shape, you must reset the height percent.

.. code-block:: php

    $oShape->getView3D()->setHeightPercent(null);

Types
-----

Area
^^^^

TODO

Bar & Bar3D
^^^^^^^^^^^

Gap Width
"""""""""

You can define the gap width between bar or columns clusters. It is defined in percent.
The default value is 150%. The value must be defined between 0 and 500.

.. code-block:: php

    $oBarChart = new Bar();
    $oBarChart->setGapWidthPercent(250);

Stacking
""""""""

You can stack multiples series in a same chart. After adding multiples series, you can define the bar grouping with `setBarGrouping` method of AbstractTypeBar.

.. code-block:: php

    $oBarChart = new Bar();
    $oBarChart->addSeries($oSeries1);
    $oBarChart->addSeries($oSeries2);
    $oBarChart->addSeries($oSeries3);
    $oBarChart->setBarGrouping(Bar::GROUPING_CLUSTERED);
    // OR
    $oBarChart->setBarGrouping(Bar::GROUPING_STACKED);
    // OR
    $oBarChart->setBarGrouping(Bar::GROUPING_PERCENTSTACKED);

- Bar::GROUPING_CLUSTERED
.. image:: images/chart_columns_52x60.png
   :width: 120px
   :alt: Bar::GROUPING_CLUSTERED

- Bar::GROUPING_STACKED
.. image:: images/chart_columnstack_52x60.png
   :width: 120px
   :alt: Bar::GROUPING_STACKED

- Bar::GROUPING_PERCENTSTACKED
.. image:: images/chart_columnpercent_52x60.png
   :width: 120px
   :alt: Bar::GROUPING_PERCENTSTACKED


Line
^^^^

TODO

Pie & Pie3D
^^^^^^^^^^^

TODO

Scatter
^^^^^^^

TODO

