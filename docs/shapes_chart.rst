.. _shapes_chart:

Charts
======

To create a chart, use `createChartShape` method of Slide.

Example:

.. code-block:: php

	$chartShape = $slide->createChartShape();

Types
-------

Area
^^^^

TODO

Bar & Bar3D
^^^^^^^^^^^

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

