# Charts

To create a chart, use `createChartShape` method of Slide.

``` php
<?php

$chartShape = $slide->createChartShape();
```

## Customization

### Manage how blank values are displayed

You can define how blank values are displayed with the method `setDisplayBlankAs`.

![Slideshow type](/images/libreoffice_chart_displayblankas.png)

Differents types are available:

* `Chart::BLANKAS_GAP` for **Leave a gap**
* `Chart::BLANKAS_ZERO` for **Assume zero** (default)
* `Chart::BLANKAS_SPAN` for **Continue line**

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart;

// Set the behavior
$chart->setDisplayBlankAs(Chart::BLANKAS_GAP);
// Get the behavior
echo $chart->getDisplayBlankAs();
```

## Parts

### Axis

#### Title

You can define title for each axis (X & Y) with `setTitle` method.
You can apply a rotation with the `setTitleRotation` method with an expected paremeter in degrees.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;

$line = new Line();

$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);

$shape->getPlotArea()->getAxisX()->setTitle('Axis X');
$shape->getPlotArea()->getAxisX()->setTitleRotation(45);
```

#### Gridlines

You can define gridlines (minor and major) for each axis (X & Y).
For each gridline, you can custom the width (in points), the fill type and the fill color.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;

$line = new Line();

$gridlines = new Gridlines();
$gridlines->getOutline()->setWidth(10);
$gridlines->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));

$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);
$shape->getPlotArea()->getAxisX()->setMajorGridlines($gridlines);
```

#### Bounds (Min & Max)

For Axis, you can define the min & max bounds with `setMinBounds` & `setMaxBounds` methods.
For resetting them, you pass null as parameter to these methods.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;

$line = new Line();

$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);
$shape->getPlotArea()->getAxisX()->setMinBounds(0);
$shape->getPlotArea()->getAxisX()->setMaxBounds(200);
```

#### Outline

You can define outline for each axis (X & Y).

``` php
<?php

$line = new Line();

$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);
$shape->getPlotArea()->getAxisX()->getOutline()->setWidth(10);
$shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
```
#### Tick Label Position

You can define the tick label position with the `setTickLabelPosition` method.
For resetting it, you pass `Axis::TICK_LABEL_POSITION_NEXT_TO` as parameter to this method.

Differents types are available:

* `Axis::TICK_LABEL_POSITION_HIGH`: **Labels are at the high end of the perpendicular axis**
* `Axis::TICK_LABEL_POSITION_LOW`: **Labels are at the low end of the perpendicular axis**
* `Axis::TICK_LABEL_POSITION_NEXT_TO`: **Labels are next to the axis** (default)

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Axis;

$line = new Line();

$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);
$shape->getPlotArea()->getAxisY()->setTickLabelPosition(Axis::TICK_LABEL_POSITION_LOW);
```
#### Tick Marks

For Axis Y, you can define tick mark with `setMinorTickMark` & `setMajorTickMark` methods.
For resetting them, you pass `Axis::TICK_MARK_NONE` as parameter to these methods.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Axis;

$line = new Line();

$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);
$shape->getPlotArea()->getAxisY()->setMinorTickMark(Axis::TICK_MARK_NONE);
$shape->getPlotArea()->getAxisY()->setMajorTickMark(Axis::TICK_MARK_INSIDE);
```

#### Unit

For Axis Y, you can define unit with `setMinorUnit` & `setMajorUnit` methods.
For resetting them, you pass null as parameter to these methods.

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Axis;

$line = new Line();

$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);
$shape->getPlotArea()->getAxisY()->setMinorUnit(null);
$shape->getPlotArea()->getAxisY()->setMajorUnit(0.05);
```
#### Visibility

You can define visibility for each axis (X & Y).

``` php
<?php

$line = new Line();

$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);
$shape->getPlotArea()->getAxisX()->setIsVisible(false);
```

### Title

By default, the title of a chart is displayed.
For hiding it, you define its visibility to false.

``` php
<?php

$line = new Line();
$shape = $slide->createChartShape();
$shape->getPlotArea()->setType($line);
// Hide the title
$shape->getTitle()->setVisible(false);
```

### Series

#### Display Informations
You can define if some informations are displayed.

``` php
<?php

$series = new Series('Downloads', $seriesData);
$series->setSeparator(';');
$series->setShowCategoryName(true);
$series->setShowLeaderLines(true);
$series->setShowLegendKey(true);
$series->setShowPercentage(true);
$series->setShowSeriesName(true);
$series->setShowValue(true);
```

#### Font
You can custom the font of a serie.

``` php
<?php

$series = new Series('Downloads', $seriesData);
// Define the size
$series->getFont()->setSize(25);
```

#### Label Position
You can define the position of the data label.
Each position is described in [MSDN](https://msdn.microsoft.com/en-us/library/mt459417(v=office.12).aspx).

``` php
<?php

$series = new Series('Downloads', $seriesData);
$series->setLabelPosition(Series::LABEL_INSIDEEND);
```

#### Marker
You can custom the marker of a serie, for Line & Scatter charts.

##### Customize the border

!!! warning
    Available only on the PowerPoint2007 Writer

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Style\Border;

$series = new Series('Downloads', $seriesData);
$marker = $series->getMarker();
$marker->getBorder()->setLineStyle(Border::LINE_SINGLE);
```

##### Customize the fill

!!! warning
    Available only on the PowerPoint2007 Writer

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Style\Fill;

$series = new Series('Downloads', $seriesData);
$marker = $series->getMarker();
$marker->getFill()->setFillType(Fill::FILL_SOLID);
```

##### Customize the symbol

``` php
<?php

use PhpOffice\PhpPresentation\Shape\Chart\Marker;

$series = new Series('Downloads', $seriesData);
$marker = $series->getMarker();
$marker->setSymbol(Marker::SYMBOL_DASH)->setSize(10);
```

#### Outline
You can custom the line of a serie, for Line & Scatter charts.

``` php
<?php

use PhpOffice\PhpPresentation\Style\Outline;

$outline = new Outline();
// Define the color
$outline->getFill()->setFillType(Fill::FILL_SOLID);
$outline->getFill()->setStartColor(new Color(Color::COLOR_YELLOW));
// Define the width (in points)
$outline->setWidth(2);

$series = new Series('Downloads', $seriesData);
$series->setOutline($outline);
```

### View3D

For enabling the autoscale for a shape, you must reset the height percent.

``` php
<?php

$shape->getView3D()->setHeightPercent(null);
```

## Types

### Area

TODO

### Bar & Bar3D

#### Gap Width

You can define the gap width between bar or columns clusters. It is defined in percent.
The default value is 150%. The value must be defined between 0 and 500.

``` php
<?php

$barChart = new Bar();
$barChart->setGapWidthPercent(250);
```

#### Stacking

You can stack multiples series in a same chart. After adding multiples series, you can define the bar grouping with `setBarGrouping` method of AbstractTypeBar.

``` php
<?php

$barChart = new Bar();
$barChart->addSeries($series1);
$barChart->addSeries($series2);
$barChart->addSeries($series3);
$barChart->setBarGrouping(Bar::GROUPING_CLUSTERED);
// OR
$barChart->setBarGrouping(Bar::GROUPING_STACKED);
// OR
$barChart->setBarGrouping(Bar::GROUPING_PERCENTSTACKED);
```

- Bar::GROUPING_CLUSTERED

![Bar::GROUPING_CLUSTERED](/images/chart_columns_52x60.png)

- Bar::GROUPING_STACKED

![Bar::GROUPING_STACKED](/images/chart_columnstack_52x60.png)

- Bar::GROUPING_PERCENTSTACKED

![Bar::GROUPING_PERCENTSTACKED](/images/chart_columnpercent_52x60.png)


### Line

TODO

### Pie & Pie3D

TODO

### Scatter

TODO

