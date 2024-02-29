.. _shapes_generic:

Generic Shape
=============

To create a dynamic shape we use `createGenericShape` method of slide.

Example:

.. code-block:: php

    $slide->createGenericShape();

Parameters
==========

-``fromX``: Starting position in terms of x.

-``fromY``: Starting point in terms of y.

-``toX``: ending x point in reference to the starting point.

-``toY``: ending y point in reference to the starting point.

-``rotation``: can be used to rotate the generated shape.

-``shape``: used to define desired shape.

Following options for ``shape`` are :

``accentBorderCallout1``
``accentBorderCallout2``
``accentBorderCallout3``
``accentCallout1``
``accentCallout2``
``accentCallout3``
``actionButtonBackPrevious``
``actionButtonBeginning``
``actionButtonBlank``
``actionButtonDocument``
``actionButtonEnd``
``actionButtonForwardNext``
``actionButtonHelp``
``actionButtonHome``
``actionButtonInformation``
``actionButtonMovie``
``actionButtonReturn``
``actionButtonSound``
``arc``
``bentArrow``
``bentConnector2``
``bentConnector3``
``bentConnector4``
``bentConnector5``
``bentUpArrow``
``bevel``
``blockArc``
``borderCallout1``
``borderCallout2``
``borderCallout3``
``bracePair``
``bracketPair``
``callout1``
``callout2``
``callout3``
``can``
``chartPlus``
``chartStar``
``chartX``
``chevron``
``chord``
``circularArrow``
``cloud``
``cloudCallout``
``corner``
``cornerTabs``
``cube``
``curvedConnector2``
``curvedConnector3``
``curvedConnector4``
``curvedConnector5``
``curvedDownArrow``
``curvedLeftArrow``
``curvedRightArrow``
``curvedUpArrow``
``decagon``
``diagStripe``
``diamond``
``dodecagon``
``donut``
``doubleWave``
``downArrow``
``downArrowCallout``
``ellipse``
``ellipseRibbon``
``ellipseRibbon2``
``flowChartAlternateProcess``
``flowChartCollate``
``flowChartConnector``
``flowChartDecision``
``flowChartDelay``
``flowChartDisplay``
``flowChartDocument``
``flowChartExtract``
``flowChartInputOutput``
``flowChartInternalStorage``
``flowChartMagneticDisk``
``flowChartMagneticDrum``
``flowChartMagneticTape``
``flowChartManualInput``
``flowChartManualOperation``
``flowChartMerge``
``flowChartMultidocument``
``flowChartOfflineStorage``
``flowChartOffpageConnector``
``flowChartOnlineStorage``
``flowChartOr``
``flowChartPredefinedProcess``
``flowChartPreparation``
``flowChartProcess``
``flowChartPunchedCard``
``flowChartPunchedTape``
``flowChartSort``
``flowChartSummingJunction``
``flowChartTerminator``
``folderCorner``
``frame``
``funnel``
``gear6``
``gear9``
``halfFrame``
``heart``
``heptagon``
``hexagon``
``homePlate``
``horizontalScroll``
``irregularSeal1``
``irregularSeal2``
``leftArrow``
``leftArrowCallout``
``leftBrace``
``leftBracket``
``leftCircularArrow``
``leftRightArrow``
``leftRightArrowCallout``
``leftRightCircularArrow``
``leftRightRibbon``
``irregularSeal1``
``leftRightUpArrow``
``leftUpArrow``
``lightningBolt``
``line``
``lineInv``
``mathDivide``
``mathEqual``
``mathMinus``
``mathMultiply``
``mathNotEqual``
``mathPlus``
``moon``
``nonIsoscelesTrapezoid``
``noSmoking``
``notchedRightArrow``
``octagon``
``parallelogram``
``pentagon``
``pie``
``pieWedge``
``plaque``
``plaqueTabs``
``plus``
``quadArrow``
``quadArrowCallout``
``rect``
``ribbon``
``ribbon2``
``rightArrow``
``rightArrowCallout``
``rightBrace``
``rightBracket``
``round1Rect``
``round2DiagRect``
``round2SameRect``
``roundRect``
``rtTriangle``
``smileyFace``
``snip1Rect``
``snip2DiagRect``
``snip2SameRect``
``snipRoundRect``
``squareTabs``
``star10``
``star12``
``star16``
``star24``
``star32``
``star4``
``star5``
``star6``
``star7``
``star8``
``straightConnector1``
``stripedRightArrow``
``sun``
``swooshArrow``
``teardrop``
``trapezoid``
``triangle``
``upArrow``
``upArrowCallout``
``upDownArrow``
``upDownArrowCallout``
``uturnArrow``
``verticalScroll``
``wave``
``wedgeEllipseCallout``
``wedgeRectCallout``
``wedgeRoundRectCallout``

Usage
=====

.. code-block:: php

    use PhpOffice\PhpPresentation\PhpPresentation;
    use PhpOffice\PhpPresentation\IOFactory;
    use PhpOffice\PhpPresentation\Style\Border;

    $objPhpPowerPoint = new PhpPresentation();
    $currentSlide = createTemplatedSlide( $objPhpPowerPoint );
    $currentSlide->createGenricShape(510,200,590,350, 0,'triangle');

Border
======

For borders we can use the following code :

.. code-block:: php

     $currentSlide->createGenricShape(310,200,390,350, 0,'triangle')->getBorder()->setLineStyle(Border::LINE_SINGLE)->setLineWidth(2)->getColor()->setARGB('FF2980b9');
