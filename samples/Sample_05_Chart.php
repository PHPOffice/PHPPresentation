<?php

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;

function fnSlide_Area(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;

    // Generate sample data for chart
    echo date('H:i:s') . ' Generate sample data for chart' . EOL;
    $seriesData = array(
        'Monday' => 12,
        'Tuesday' => 15,
        'Wednesday' => 13,
        'Thursday' => 17,
        'Friday' => 14,
        'Saturday' => 9,
        'Sunday' => 7
    );

    // Create templated slide
    echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Create a line chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a area chart (that should be inserted in a chart shape)' . EOL;
    $areaChart = new Area();
    $series = new Series('Downloads', $seriesData);
    $series->setShowSeriesName(true);
    $series->setShowValue(true);
    $series->getFill()->setStartColor(new Color('FF93A9CE'));
    $series->setLabelPosition(Series::LABEL_INSIDEEND);
    $areaChart->addSeries($series);

    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)' . EOL;
    $shape = $currentSlide->createChartShape();
    $shape->getTitle()->setVisible(false);
    $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
    $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Daily Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getPlotArea()->setType($areaChart);
    $shape->getPlotArea()->getAxisX()->setTitle('Axis X');
    $shape->getPlotArea()->getAxisY()->setTitle('Axis Y');
    $shape->getView3D()->setRotationX(30);
    $shape->getView3D()->setPerspective(30);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}

function fnSlide_Bar(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;

    // Create templated slide
    echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Generate sample data for first chart
    echo date('H:i:s') . ' Generate sample data for chart'.EOL;
    $series1Data = array('Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293);
    $series2Data = array('Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379);

    // Create a bar chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a bar chart (that should be inserted in a chart shape)'.EOL;
    $barChart = new Bar();
    $barChart->setGapWidthPercent(158);
    $series1 = new Series('2009', $series1Data);
    $series1->setShowSeriesName(true);
    $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4F81BD'));
    $series1->getFont()->getColor()->setRGB('00FF00');
    $series1->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
    $series2 = new Series('2010', $series2Data);
    $series2->setShowSeriesName(true);
    $series2->getFont()->getColor()->setRGB('FF0000');
    $series2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFC0504D'));
    $series2->setLabelPosition(Series::LABEL_INSIDEEND);
    $barChart->addSeries($series1);
    $barChart->addSeries($series2);

    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)'.EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Monthly Downloads')
        ->setResizeProportional(false)
        ->setHeight(550)
        ->setWidth(700)
        ->setOffsetX(120)
        ->setOffsetY(80);
    $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Monthly Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    $shape->getPlotArea()->getAxisX()->setTitle('Month');
    $shape->getPlotArea()->getAxisY()->getFont()->getColor()->setRGB('00FF00');
    $shape->getPlotArea()->getAxisY()->setTitle('Downloads');
    $shape->getPlotArea()->setType($barChart);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}

function fnSlide_BarHorizontal(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;

    // Create a bar chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a horizontal bar chart (that should be inserted in a chart shape) '.EOL;
    $barChartHorz = clone $objPHPPresentation->getSlide(1)->getShapeCollection()->offsetGet(1)->getPlotArea()->getType();
    $barChartHorz->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);

    // Create templated slide
    echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)'.EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Monthly Downloads')
        ->setResizeProportional(false)
        ->setHeight(550)
        ->setWidth(700)
        ->setOffsetX(120)
        ->setOffsetY(80);
    $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Monthly Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    $shape->getPlotArea()->getAxisX()->setTitle('Month');
    $shape->getPlotArea()->getAxisY()->setTitle('Downloads');
    $shape->getPlotArea()->setType($barChartHorz);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}

function fnSlide_BarStacked(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;

    // Create templated slide
    echo EOL . date( 'H:i:s' ) . ' Create templated slide' . EOL;
    $currentSlide = createTemplatedSlide( $objPHPPresentation );

    // Generate sample data for first chart
    echo date( 'H:i:s' ) . ' Generate sample data for chart' . EOL;
    $series1Data = array('Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293);
    $series2Data = array('Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379);
    $series3Data = array('Jan' => 233, 'Feb' => 146, 'Mar' => 238, 'Apr' => 175, 'May' => 108, 'Jun' => 257, 'Jul' => 199, 'Aug' => 201, 'Sep' => 88, 'Oct' => 147, 'Nov' => 287, 'Dec' => 105);

    // Create a bar chart (that should be inserted in a shape)
    echo date( 'H:i:s' ) . ' Create a stacked bar chart (that should be inserted in a chart shape)' . EOL;
    $StackedBarChart = new Bar();
    $series1 = new Series( '2009', $series1Data );
    $series1->setShowSeriesName( false );
    $series1->getFill()->setFillType( Fill::FILL_SOLID )->setStartColor( new Color( 'FF4F81BD' ) );
    $series1->getFont()->getColor()->setRGB( '00FF00' );
    $series1->setShowValue( true );
    $series1->setShowPercentage( false );
    $series2 = new Series( '2010', $series2Data );
    $series2->setShowSeriesName( false );
    $series2->getFont()->getColor()->setRGB( 'FF0000' );
    $series2->getFill()->setFillType( Fill::FILL_SOLID )->setStartColor( new Color( 'FFC0504D' ) );
    $series2->setShowValue( true );
    $series2->setShowPercentage( false );
    $series3 = new Series( '2011', $series3Data );
    $series3->setShowSeriesName( false );
    $series3->getFont()->getColor()->setRGB( 'FF0000' );
    $series3->getFill()->setFillType( Fill::FILL_SOLID )->setStartColor( new Color( 'FF804DC0' ) );
    $series3->setShowValue( true );
    $series3->setShowPercentage( false );
    $StackedBarChart->addSeries( $series1 );
    $StackedBarChart->addSeries( $series2 );
    $StackedBarChart->addSeries( $series3 );
    $StackedBarChart->setBarGrouping( Bar::GROUPING_STACKED );
    // Create a shape (chart)
    echo date( 'H:i:s' ) . ' Create a shape (chart)' . EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName( 'PHPPresentation Monthly Downloads' )
        ->setResizeProportional( false )
        ->setHeight( 550 )
        ->setWidth( 700 )
        ->setOffsetX( 120 )
        ->setOffsetY( 80 );
    $shape->setShadow( $oShadow );
    $shape->setFill( $oFill );
    $shape->getBorder()->setLineStyle( Border::LINE_SINGLE );
    $shape->getTitle()->setText( 'PHPPresentation Monthly Downloads' );
    $shape->getTitle()->getFont()->setItalic( true );
    $shape->getTitle()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_RIGHT );
    $shape->getPlotArea()->getAxisX()->setTitle( 'Month' );
    $shape->getPlotArea()->getAxisY()->setTitle( 'Downloads' );
    $shape->getPlotArea()->setType( $StackedBarChart );
    $shape->getLegend()->getBorder()->setLineStyle( Border::LINE_SINGLE );
    $shape->getLegend()->getFont()->setItalic( true );
}

function fnSlide_BarPercentStacked(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;

    // Create templated slide
    echo EOL . date( 'H:i:s' ) . ' Create templated slide' . EOL;
    $currentSlide = createTemplatedSlide( $objPHPPresentation );

    // Generate sample data for first chart
    echo date( 'H:i:s' ) . ' Generate sample data for chart' . EOL;
    $series1Data = array('Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293);
    $Series1Sum = array_sum($series1Data);
    foreach ($series1Data as $CatName => $Value) {
        $series1Data[$CatName]= round($Value / $Series1Sum, 2);
    }
    $series2Data = array('Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379);
    $Series2Sum = array_sum($series2Data);
    foreach ($series2Data as $CatName => $Value) {
        $series2Data[$CatName] = round($Value / $Series2Sum, 2);
    }
    $series3Data = array('Jan' => 233, 'Feb' => 146, 'Mar' => 238, 'Apr' => 175, 'May' => 108, 'Jun' => 257, 'Jul' => 199, 'Aug' => 201, 'Sep' => 88, 'Oct' => 147, 'Nov' => 287, 'Dec' => 105);
    $Series3Sum = array_sum( $series3Data );
    foreach ($series3Data as $CatName => $Value) {
        $series3Data[$CatName] = round($Value / $Series3Sum,2);
    }

    // Create a bar chart (that should be inserted in a shape)
    echo date( 'H:i:s' ) . ' Create a percent stacked horizontal bar chart (that should be inserted in a chart shape)' . EOL;
    $PercentStackedBarChartHoriz = new Bar();
    $series1 = new Series( '2009', $series1Data );
    $series1->setShowSeriesName( false );
    $series1->getFill()->setFillType( Fill::FILL_SOLID )->setStartColor( new Color( 'FF4F81BD' ) );
    $series1->getFont()->getColor()->setRGB( '00FF00' );
    $series1->setShowValue( true );
    $series1->setShowPercentage( false );
    // Set Data Label Format For Chart To Display Percent
    $series1->setDlblNumFormat( '#%' );
    $series2 = new Series( '2010', $series2Data );
    $series2->setShowSeriesName( false );
    $series2->getFont()->getColor()->setRGB( 'FF0000' );
    $series2->getFill()->setFillType( Fill::FILL_SOLID )->setStartColor( new Color( 'FFC0504D' ) );
    $series2->setShowValue( true );
    $series2->setShowPercentage( false );
    $series2->setDlblNumFormat( '#%' );
    $series3 = new Series( '2011', $series3Data );
    $series3->setShowSeriesName( false );
    $series3->getFont()->getColor()->setRGB( 'FF0000' );
    $series3->getFill()->setFillType( Fill::FILL_SOLID )->setStartColor( new Color( 'FF804DC0' ) );
    $series3->setShowValue( true );
    $series3->setShowPercentage( false );
    $series3->setDlblNumFormat( '#%' );
    $PercentStackedBarChartHoriz->addSeries( $series1 );
    $PercentStackedBarChartHoriz->addSeries( $series2 );
    $PercentStackedBarChartHoriz->addSeries( $series3 );
    $PercentStackedBarChartHoriz->setBarGrouping( Bar::GROUPING_PERCENTSTACKED );
    $PercentStackedBarChartHoriz->setBarDirection( Bar3D::DIRECTION_HORIZONTAL );
    // Create a shape (chart)
    echo date( 'H:i:s' ) . ' Create a shape (chart)' . EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName( 'PHPPresentation Monthly Downloads' )
        ->setResizeProportional( false )
        ->setHeight( 550 )
        ->setWidth( 700 )
        ->setOffsetX( 120 )
        ->setOffsetY( 80 );
    $shape->setShadow( $oShadow );
    $shape->setFill( $oFill );
    $shape->getBorder()->setLineStyle( Border::LINE_SINGLE );
    $shape->getTitle()->setText( 'PHPPresentation Monthly Downloads' );
    $shape->getTitle()->getFont()->setItalic( true );
    $shape->getTitle()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_RIGHT );
    $shape->getPlotArea()->getAxisX()->setTitle( 'Month' );
    $shape->getPlotArea()->getAxisY()->setTitle( 'Downloads' );
    $shape->getPlotArea()->setType( $PercentStackedBarChartHoriz );
    $shape->getLegend()->getBorder()->setLineStyle( Border::LINE_SINGLE );
    $shape->getLegend()->getFont()->setItalic( true );
}

function fnSlide_Bar3D(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;

    // Create templated slide
    echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Generate sample data for first chart
    echo date('H:i:s') . ' Generate sample data for chart'.EOL;
    $series1Data = array('Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293);
    $series2Data = array('Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379);

    // Create a bar chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a bar chart (that should be inserted in a chart shape)'.EOL;
    $bar3DChart = new Bar3D();
    $series1 = new Series('2009', $series1Data);
    $series1->setShowSeriesName(true);
    $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4F81BD'));
    $series1->getFont()->getColor()->setRGB('00FF00');
    $series1->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
    $series2 = new Series('2010', $series2Data);
    $series2->setShowSeriesName(true);
    $series2->getFont()->getColor()->setRGB('FF0000');
    $series2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFC0504D'));
    $bar3DChart->addSeries($series1);
    $bar3DChart->addSeries($series2);

    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)'.EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Monthly Downloads')
        ->setResizeProportional(false)
        ->setHeight(550)
        ->setWidth(700)
        ->setOffsetX(120)
        ->setOffsetY(80);
    $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Monthly Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    $shape->getPlotArea()->getAxisX()->setTitle('Month');
    $shape->getPlotArea()->getAxisY()->setTitle('Downloads');
    $shape->getPlotArea()->setType($bar3DChart);
    $shape->getView3D()->setRightAngleAxes(true);
    $shape->getView3D()->setRotationX(20);
    $shape->getView3D()->setRotationY(20);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}

function fnSlide_Bar3DHorizontal(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;
    
    // Create a bar chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a horizontal bar chart (that should be inserted in a chart shape) '.EOL;
    $bar3DChartHorz = clone $objPHPPresentation->getSlide(5)->getShapeCollection()->offsetGet(1)->getPlotArea()->getType();
    $bar3DChartHorz->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);
    
    // Create templated slide
    echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);
    
    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)'.EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Monthly Downloads')
            ->setResizeProportional(false)
            ->setHeight(550)
            ->setWidth(700)
            ->setOffsetX(120)
            ->setOffsetY(80);
    $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Monthly Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    $shape->getPlotArea()->getAxisX()->setTitle('Month');
    $shape->getPlotArea()->getAxisY()->setTitle('Downloads');
    $shape->getPlotArea()->setType($bar3DChartHorz);
    $shape->getView3D()->setRightAngleAxes(true);
    $shape->getView3D()->setRotationX(20);
    $shape->getView3D()->setRotationY(20);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}

function fnSlide_Pie3D(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;
    
    // Create templated slide
    echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);
    
    // Generate sample data for second chart
    echo date('H:i:s') . ' Generate sample data for chart'.EOL;
    $seriesData = array('Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7);
    
    // Create a pie chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a pie chart (that should be inserted in a chart shape)'.EOL;
    $pie3DChart = new Pie3D();
    $pie3DChart->setExplosion(20);
    $series = new Series('Downloads', $seriesData);
    $series->setShowSeriesName(true);
    $series->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
    $series->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
    $series->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
    $series->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF725990'));
    $series->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4299B0'));
    $series->getDataPointFill(5)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFDC853E'));
    $series->getDataPointFill(6)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF93A9CE'));
    $pie3DChart->addSeries($series);

    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)'.EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Daily Downloads')
          ->setResizeProportional(false)
          ->setHeight(550)
          ->setWidth(700)
          ->setOffsetX(120)
          ->setOffsetY(80);
    $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Daily Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getPlotArea()->setType($pie3DChart);
    $shape->getView3D()->setRotationX(30);
    $shape->getView3D()->setPerspective(30);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}

function fnSlide_Pie(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;

    // Create templated slide
    echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Generate sample data for second chart
    echo date('H:i:s') . ' Generate sample data for chart'.EOL;
    $seriesData = array('Monday' => 18, 'Tuesday' => 23, 'Wednesday' => 14, 'Thursday' => 12, 'Friday' => 20, 'Saturday' => 8, 'Sunday' => 10);

    // Create a pie chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a non-3D pie chart (that should be inserted in a chart shape)'.EOL;
    $pieChart = new Pie();
    $pieChart->setExplosion(15);
    $series = new Series('Downloads', $seriesData);
    $series->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF7CB5EC'));
    $series->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF434348'));
    $series->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF90ED7D'));
    $series->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFF7A35C'));
    $series->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8085E9'));
    $series->getDataPointFill(5)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFF15C80'));
    $series->getDataPointFill(6)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE4D354'));
    $series->setShowPercentage( true );
    $series->setShowValue( false );
    $series->setShowSeriesName( false );
    $series->setShowCategoryName( true );
    $series->setDlblNumFormat('%d');
    $pieChart->addSeries($series);

    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)'.EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Daily Downloads')
          ->setResizeProportional(false)
          ->setHeight(550)
          ->setWidth(700)
          ->setOffsetX(120)
          ->setOffsetY(80);
    $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Daily Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getPlotArea()->setType($pieChart);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}

function fnSlide_Scatter(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;
    
    // Create templated slide
    echo EOL.date('H:i:s') . ' Create templated slide'.EOL;
    $currentSlide = createTemplatedSlide($objPHPPresentation); // local function
    
    // Generate sample data for fourth chart
    echo date('H:i:s') . ' Generate sample data for chart'.EOL;
    $seriesData = array('Monday' => 0.1, 'Tuesday' => 0.33333, 'Wednesday' => 0.4444, 'Thursday' => 0.5, 'Friday' => 0.4666, 'Saturday' => 0.3666, 'Sunday' => 0.1666);
    
    // Create a scatter chart (that should be inserted in a shape)
    echo date('H:i:s') . ' Create a scatter chart (that should be inserted in a chart shape)'.EOL;
    $lineChart = new Scatter();
    $series = new Series('Downloads', $seriesData);
    $series->setShowSeriesName(true);
    $series->getMarker()->setSymbol(\PhpOffice\PhpPresentation\Shape\Chart\Marker::SYMBOL_DASH);
    $series->getMarker()->setSize(10);
    $lineChart->addSeries($series);
    
    // Create a shape (chart)
    echo date('H:i:s') . ' Create a shape (chart)'.EOL;
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Daily Download Distribution')
        ->setResizeProportional(false)
        ->setHeight(550)
        ->setWidth(700)
        ->setOffsetX(120)
        ->setOffsetY(80);
    $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Daily Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getPlotArea()->setType($lineChart);
    $shape->getView3D()->setRotationX(30);
    $shape->getView3D()->setPerspective(30);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}

// Create new PHPPresentation object
echo date('H:i:s') . ' Create new PHPPresentation object'.EOL;
$objPHPPresentation = new PhpPresentation();

// Set properties
echo date('H:i:s') . ' Set properties'.EOL;
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPresentation Team')
                                  ->setTitle('Sample 07 Title')
                                  ->setSubject('Sample 07 Subject')
                                  ->setDescription('Sample 07 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Remove first slide
echo date('H:i:s') . ' Remove first slide'.EOL;
$objPHPPresentation->removeSlideByIndex(0);

// Set Style
$oFill = new Fill();
$oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

$oShadow = new Shadow();
$oShadow->setVisible(true)->setDirection(45)->setDistance(10);

fnSlide_Area($objPHPPresentation);

fnSlide_Bar($objPHPPresentation);

fnSlide_BarStacked($objPHPPresentation);

fnSlide_BarPercentStacked($objPHPPresentation);

fnSlide_BarHorizontal($objPHPPresentation);

fnSlide_Bar3D($objPHPPresentation);

fnSlide_Bar3DHorizontal($objPHPPresentation);

fnSlide_Pie3D($objPHPPresentation);

fnSlide_Pie($objPHPPresentation);

fnSlide_Scatter($objPHPPresentation);

// Save file
echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
    include_once 'Sample_Footer.php';
}
