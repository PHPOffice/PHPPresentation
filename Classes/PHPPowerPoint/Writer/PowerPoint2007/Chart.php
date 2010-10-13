<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Writer_PowerPoint2007_Chart
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_Chart extends PHPPowerPoint_Writer_PowerPoint2007_Slide
{	
	/**
	 * Write chart to XML format
	 *
	 * @param	PHPPowerPoint_Shape_Chart	$chart
	 * @return	string						XML Output
	 * @throws	Exception
	 */
	public function writeChart(PHPPowerPoint_Shape_Chart $chart = null)
	{
		// Check slide
		if (is_null($chart))
			throw new Exception("Invalid PHPPowerPoint_Shape_Chart object passed.");

		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// c:chartSpace
		$objWriter->startElement('c:chartSpace');
		$objWriter->writeAttribute('xmlns:c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
		$objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
		$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		
			// c:date1904
			$objWriter->startElement('c:date1904');
			$objWriter->writeAttribute('val', '1');
			$objWriter->endElement();
			
			// c:lang
			$objWriter->startElement('c:lang');
			$objWriter->writeAttribute('val', 'en-US');
			$objWriter->endElement();
			
			// c:chart
			$objWriter->startElement('c:chart');
			
				// Title?
				if ($chart->getTitle()->getVisible()) {
					// Write title
					$this->_writeTitle($objWriter, $chart->getTitle());
				}
								
				// c:autoTitleDeleted
				$objWriter->startElement('c:autoTitleDeleted');
				$objWriter->writeAttribute('val', '0');
				$objWriter->endElement();
				
				// c:view3D
				$objWriter->startElement('c:view3D');
				
					// c:rotX
					$objWriter->startElement('c:rotX');
					$objWriter->writeAttribute('val', $chart->getView3D()->getRotationX());
					$objWriter->endElement();
					
					// c:hPercent
					$objWriter->startElement('c:hPercent');
					$objWriter->writeAttribute('val', $chart->getView3D()->getHeightPercent());
					$objWriter->endElement();
					
					// c:rotY
					$objWriter->startElement('c:rotY');
					$objWriter->writeAttribute('val', $chart->getView3D()->getRotationY());
					$objWriter->endElement();
					
					// c:depthPercent
					$objWriter->startElement('c:depthPercent');
					$objWriter->writeAttribute('val', $chart->getView3D()->getDepthPercent());
					$objWriter->endElement();
					
					// c:rAngAx
					$objWriter->startElement('c:rAngAx');
					$objWriter->writeAttribute('val', $chart->getView3D()->getRightAngleAxes() ? '1' : '0');
					$objWriter->endElement();
					
					// c:perspective
					$objWriter->startElement('c:perspective');
					$objWriter->writeAttribute('val', $chart->getView3D()->getPerspective());
					$objWriter->endElement();
    
				$objWriter->endElement();
				
				// Write plot area
				$this->_writePlotArea($objWriter, $chart->getPlotArea());
    
				// Legend?
				if ($chart->getLegend()->getVisible()) {
					// Write legend
					$this->_writeLegend($objWriter, $chart->getLegend());
				}
				
				// c:plotVisOnly
				$objWriter->startElement('c:plotVisOnly');
				$objWriter->writeAttribute('val', '1');
				$objWriter->endElement();

			$objWriter->endElement();
			
			// c:spPr
			$objWriter->startElement('c:spPr');
			
				// Fill
				$this->_writeFill($objWriter, $chart->getFill());
							
				// Border
				if ($chart->getBorder()->getLineStyle() != PHPPowerPoint_Style_Border::LINE_NONE) {
					$this->_writeBorder($objWriter,$chart->getBorder(),'');
				}
				
				// Shadow
				if ($chart->getShadow()->getVisible()) {
					// a:effectLst
					$objWriter->startElement('a:effectLst');

						// a:outerShdw
						$objWriter->startElement('a:outerShdw');
						$objWriter->writeAttribute('blurRad', 		PHPPowerPoint_Shared_Drawing::pixelsToEMU($chart->getShadow()->getBlurRadius()));
						$objWriter->writeAttribute('dist',			PHPPowerPoint_Shared_Drawing::pixelsToEMU($chart->getShadow()->getDistance()));
						$objWriter->writeAttribute('dir',			PHPPowerPoint_Shared_Drawing::degreesToAngle($chart->getShadow()->getDirection()));
						$objWriter->writeAttribute('algn',			$chart->getShadow()->getAlignment());
						$objWriter->writeAttribute('rotWithShape', 	'0');

							// a:srgbClr
							$objWriter->startElement('a:srgbClr');
							$objWriter->writeAttribute('val',		$chart->getShadow()->getColor()->getRGB());

								// a:alpha
								$objWriter->startElement('a:alpha');
								$objWriter->writeAttribute('val', 	$chart->getShadow()->getAlpha() * 1000);
								$objWriter->endElement();

							$objWriter->endElement();

						$objWriter->endElement();

					$objWriter->endElement();
				}
				
			$objWriter->endElement();
		
		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}
	
	/**
	 * Write Title
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Shape_Chart_Title	$subject
	 * @throws 	Exception
	 */
	protected function _writeTitle(PHPPowerPoint_Shared_XMLWriter $objWriter, PHPPowerPoint_Shape_Chart_Title $subject) {
		// c:title
		$objWriter->startElement('c:title');
					
			// c:tx
			$objWriter->startElement('c:tx');
						
				// c:rich
				$objWriter->startElement('c:rich');
							
					// a:bodyPr
					$objWriter->writeElement('a:bodyPr', null);
								
					// a:lstStyle
					$objWriter->writeElement('a:lstStyle', null);
								
					// a:p
					$objWriter->startElement('a:p');
								
						// a:pPr
						$objWriter->startElement('a:pPr');
									
							// a:defRPr
							$objWriter->writeElement('a:defRPr', null);
									
						$objWriter->endElement();
									
						// a:r
						$objWriter->startElement('a:r');
									
							// a:rPr
							$objWriter->startElement('a:rPr');
							$objWriter->writeAttribute('lang', 'en-US');
							$objWriter->writeAttribute('dirty', '0');
							$objWriter->endElement();
										
							// a:t
							$objWriter->writeElement('a:t', $subject->getText());
																		
						$objWriter->endElement();
								
						// a:endParaRPr
						$objWriter->startElement('a:endParaRPr');
						$objWriter->writeAttribute('lang', 'en-US');
						$objWriter->writeAttribute('dirty', '0');
						$objWriter->endElement();
									
					$objWriter->endElement();
								
				$objWriter->endElement();
							
			$objWriter->endElement();
						
			// Write layout
			$this->_writeLayout($objWriter, $subject);
						
			// c:overlay
			$objWriter->startElement('c:overlay');
			$objWriter->writeAttribute('val', '0');
			$objWriter->endElement();
					
		$objWriter->endElement();
	}

	/**
	 * Write Plot Area
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 		$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Shape_Chart_PlotArea	$subject
	 * @throws 	Exception
	 */
	protected function _writePlotArea(PHPPowerPoint_Shared_XMLWriter $objWriter, PHPPowerPoint_Shape_Chart_PlotArea $subject) {
		// c:plotArea
		$objWriter->startElement('c:plotArea');
					
			// Write layout
			$this->_writeLayout($objWriter, $subject);

			// Write chart
			$chartType = $subject->getType();
			if ($chartType instanceof PHPPowerPoint_Shape_Chart_Type_Bar3D) {
				$this->_writeTypeBar3D($objWriter, $chartType);
			} else if ($chartType instanceof PHPPowerPoint_Shape_Chart_Type_Pie3D) {
				$this->_writeTypePie3D($objWriter, $chartType);
			} else {
				throw new Exception('The chart type provided could not be rendered.');
			}	
	
			// Write X axis?
			if ($chartType->hasAxisX()) {
				// c:catAx (Axis X)
				$objWriter->startElement('c:catAx');
				
					// c:axId
					$objWriter->startElement('c:axId');
					$objWriter->writeAttribute('val', '52743552');
					$objWriter->endElement();
					
					// c:scaling
					$objWriter->startElement('c:scaling');
					
						// c:orientation
						$objWriter->startElement('c:orientation');
						$objWriter->writeAttribute('val', 'minMax');
						$objWriter->endElement();
					
					$objWriter->endElement();
					
					// c:axPos
					$objWriter->startElement('c:axPos');
					$objWriter->writeAttribute('val', 'b');
					$objWriter->endElement();
					
					// c:numFmt
					$objWriter->startElement('c:numFmt');
					$objWriter->writeAttribute('formatCode', $subject->getAxisX()->getFormatCode());
					$objWriter->writeAttribute('sourceLinked', '0');
					$objWriter->endElement();
					
					// c:majorTickMark
					$objWriter->startElement('c:majorTickMark');
					$objWriter->writeAttribute('val', 'none');
					$objWriter->endElement();
					
					// c:tickLblPos
					$objWriter->startElement('c:tickLblPos');
					$objWriter->writeAttribute('val', 'nextTo');
					$objWriter->endElement();
	        
					// c:txPr
					$objWriter->startElement('c:txPr');
									
						// a:bodyPr
						$objWriter->writeElement('a:bodyPr', null);
									
						// a:lstStyle
						$objWriter->writeElement('a:lstStyle', null);
									
						// a:p
						$objWriter->startElement('a:p');
									
							// a:pPr
							$objWriter->startElement('a:pPr');
										
								// a:defRPr
								$objWriter->writeElement('a:defRPr', null);
										
							$objWriter->endElement();
										
							// a:r
							$objWriter->startElement('a:r');
										
								// a:rPr
								$objWriter->startElement('a:rPr');
								$objWriter->writeAttribute('lang', 'en-US');
								$objWriter->writeAttribute('dirty', '0');
								$objWriter->endElement();
											
								// a:t
								$objWriter->writeElement('a:t', $subject->getAxisX()->getTitle());
												
							$objWriter->endElement();	
							
							// a:endParaRPr
							$objWriter->startElement('a:endParaRPr');
							$objWriter->writeAttribute('lang', 'en-US');
							$objWriter->writeAttribute('dirty', '0');
							$objWriter->endElement();
										
						$objWriter->endElement();
									
					$objWriter->endElement();
								
					// c:crossAx
					$objWriter->startElement('c:crossAx');
					$objWriter->writeAttribute('val', '52749440');
					$objWriter->endElement();
					
					// c:crosses
					$objWriter->startElement('c:crosses');
					$objWriter->writeAttribute('val', 'autoZero');
					$objWriter->endElement();
					
					// c:lblAlgn
					$objWriter->startElement('c:lblAlgn');
					$objWriter->writeAttribute('val', 'ctr');
					$objWriter->endElement();
					
					// c:lblOffset
					$objWriter->startElement('c:lblOffset');
					$objWriter->writeAttribute('val', '100');
					$objWriter->endElement();
				
				$objWriter->endElement();
			}
			
			// Write Y axis?
			if ($chartType->hasAxisY()) {
				// c:valAx (Axis Y)
				$objWriter->startElement('c:valAx');
				
					// c:axId
					$objWriter->startElement('c:axId');
					$objWriter->writeAttribute('val', '52749440');
					$objWriter->endElement();		
					
					// c:scaling
					$objWriter->startElement('c:scaling');
					
						// c:orientation
						$objWriter->startElement('c:orientation');
						$objWriter->writeAttribute('val', 'minMax');
						$objWriter->endElement();
					
					$objWriter->endElement();
					
					// c:axPos
					$objWriter->startElement('c:axPos');
					$objWriter->writeAttribute('val', 'l');
					$objWriter->endElement();
					
					// c:numFmt
					$objWriter->startElement('c:numFmt');
					$objWriter->writeAttribute('formatCode', $subject->getAxisY()->getFormatCode());
					$objWriter->writeAttribute('sourceLinked', '0');
					$objWriter->endElement();
					
					// c:majorTickMark
					$objWriter->startElement('c:majorTickMark');
					$objWriter->writeAttribute('val', 'none');
					$objWriter->endElement();
					
					// c:tickLblPos
					$objWriter->startElement('c:tickLblPos');
					$objWriter->writeAttribute('val', 'nextTo');
					$objWriter->endElement();
	
					// c:txPr
					$objWriter->startElement('c:txPr');
									
						// a:bodyPr
						$objWriter->writeElement('a:bodyPr', null);
									
						// a:lstStyle
						$objWriter->writeElement('a:lstStyle', null);
									
						// a:p
						$objWriter->startElement('a:p');
									
							// a:pPr
							$objWriter->startElement('a:pPr');
										
								// a:defRPr
								$objWriter->writeElement('a:defRPr', null);
										
							$objWriter->endElement();
										
							// a:r
							$objWriter->startElement('a:r');
										
								// a:rPr
								$objWriter->startElement('a:rPr');
								$objWriter->writeAttribute('lang', 'en-US');
								$objWriter->writeAttribute('dirty', '0');
								$objWriter->endElement();
											
								// a:t
								$objWriter->writeElement('a:t', $subject->getAxisY()->getTitle());
													
							$objWriter->endElement();	
							
							// a:endParaRPr
							$objWriter->startElement('a:endParaRPr');
							$objWriter->writeAttribute('lang', 'en-US');
							$objWriter->writeAttribute('dirty', '0');
							$objWriter->endElement();
										
						$objWriter->endElement();
									
					$objWriter->endElement();
			
					// c:crossAx
					$objWriter->startElement('c:crossAx');
					$objWriter->writeAttribute('val', '52743552');
					$objWriter->endElement();
					
					// c:crosses
					$objWriter->startElement('c:crosses');
					$objWriter->writeAttribute('val', 'autoZero');
					$objWriter->endElement();
					
					// c:crossBetween
					$objWriter->startElement('c:crossBetween');
					$objWriter->writeAttribute('val', 'between');
					$objWriter->endElement();
				
				$objWriter->endElement();
			}
			
		$objWriter->endElement();
	}
	
	/**
	 * Write Legend
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 		$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Shape_Chart_Legend	$subject
	 * @throws 	Exception
	 */
	protected function _writeLegend(PHPPowerPoint_Shared_XMLWriter $objWriter, PHPPowerPoint_Shape_Chart_Legend $subject) {
		// c:legend
		$objWriter->startElement('c:legend');
					
			// c:legendPos
			$objWriter->startElement('c:legendPos');
			$objWriter->writeAttribute('val', $subject->getPosition());
			$objWriter->endElement();

			// Write layout
			$this->_writeLayout($objWriter, $subject);
						
			// c:overlay
			$objWriter->startElement('c:overlay');
			$objWriter->writeAttribute('val', '0');
			$objWriter->endElement();
					
		$objWriter->endElement();
	}
	
	/**
	 * Write Layout
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	mixed	$subject
	 * @throws 	Exception
	 */
	protected function _writeLayout(PHPPowerPoint_Shared_XMLWriter $objWriter, $subject) {
		// c:layout
		$objWriter->startElement('c:layout');

			// c:manualLayout
			$objWriter->startElement('c:manualLayout');
				// c:xMode
				$objWriter->startElement('c:xMode');
				$objWriter->writeAttribute('val', 'edge');
				$objWriter->endElement();
								
				// c:yMode
				$objWriter->startElement('c:yMode');
				$objWriter->writeAttribute('val', 'edge');
				$objWriter->endElement();

				if ($subject->getOffsetX() != 0) {
					// c:x
					$objWriter->startElement('c:x');
					$objWriter->writeAttribute('val', $subject->getOffsetX());
					$objWriter->endElement();
				}
				
				if ($subject->getOffsetY() != 0) {			
					// c:y
					$objWriter->startElement('c:y');
					$objWriter->writeAttribute('val', $subject->getOffsetY());
					$objWriter->endElement();
				}
				
				if ($subject->getWidth() != 0) {			
					// c:w
					$objWriter->startElement('c:w');
					$objWriter->writeAttribute('val', $subject->getWidth());
					$objWriter->endElement();
				}
				
				if ($subject->getHeight() != 0) {			
					// c:h
					$objWriter->startElement('c:h');
					$objWriter->writeAttribute('val', $subject->getHeight());
					$objWriter->endElement();
				}
							
			$objWriter->endElement();
						
		$objWriter->endElement();
	}
	
	/**
	 * Write Type Bar3D
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 			$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Shape_Chart_Type_Bar3D	$subject
	 * @throws 	Exception
	 */
	protected function _writeTypeBar3D(PHPPowerPoint_Shared_XMLWriter $objWriter, PHPPowerPoint_Shape_Chart_Type_Bar3D $subject) {
		// c:bar3DChart
		$objWriter->startElement('c:bar3DChart');
		
			// c:barDir
			$objWriter->startElement('c:barDir');
			$objWriter->writeAttribute('val', 'col');
			$objWriter->endElement();
			
			// c:grouping
			$objWriter->startElement('c:grouping');
			$objWriter->writeAttribute('val', 'clustered');
			$objWriter->endElement();
			
			// Write series
			$seriesIndex = 0;
			foreach ($subject->getData() as $series) {
				// c:ser
				$objWriter->startElement('c:ser');
				
					// c:idx
					$objWriter->startElement('c:idx');
					$objWriter->writeAttribute('val', $seriesIndex);
					$objWriter->endElement();
					
					// c:order
					$objWriter->startElement('c:order');
					$objWriter->writeAttribute('val', $seriesIndex);
					$objWriter->endElement();
					
					// c:tx
					$objWriter->startElement('c:tx');
					$objWriter->writeElement('c:v', $series->getTitle());
					$objWriter->endElement();
					
					// c:spPr
					$objWriter->startElement('c:spPr');
					
						// Write fill
						$this->_writeFill($objWriter, $series->getFill());
					
					$objWriter->endElement();
					
					// Write X axis data
					$axisXData = array_keys($series->getValues());
					
					// c:cat
					$objWriter->startElement('c:cat');
					
						// c:strLit / c:numLit
						if (is_int($axisXData[0]) || is_float($axisXData[0])) {
							$objWriter->startElement('c:numLit');
						} else {
							$objWriter->startElement('c:strLit');
						}
						
							// c:ptCount
							$objWriter->startElement('c:ptCount');
							$objWriter->writeAttribute('val', count($axisXData));
							$objWriter->endElement();
							
							// Add points
							for ($i = 0; $i < count($axisXData); $i++) {
								// c:pt
								$objWriter->startElement('c:pt');
								$objWriter->writeAttribute('idx', $i);
								$objWriter->writeElement('c:v', $axisXData[$i]);
								$objWriter->endElement();
							}
			
						$objWriter->endElement();
					
					$objWriter->endElement();
					
					// Write Y axis data
					$axisYData = array_values($series->getValues());
					
					// c:val
					$objWriter->startElement('c:val');
					
						// c:strLit / c:numLit
						if (is_int($axisYData[0]) || is_float($axisYData[0])) {
							$objWriter->startElement('c:numLit');
						} else {
							$objWriter->startElement('c:strLit');
						}
						
							// c:ptCount
							$objWriter->startElement('c:ptCount');
							$objWriter->writeAttribute('val', count($axisYData));
							$objWriter->endElement();
							
							// Add points
							for ($i = 0; $i < count($axisYData); $i++) {
								// c:pt
								$objWriter->startElement('c:pt');
								$objWriter->writeAttribute('idx', $i);
								$objWriter->writeElement('c:v', $axisYData[$i]);
								$objWriter->endElement();
							}
			
						$objWriter->endElement();
					
					$objWriter->endElement();
				
				$objWriter->endElement();
				
				++$seriesIndex;
			}

			// c:dLbls
			$objWriter->startElement('c:dLbls');
			
				// c:showVal
				$objWriter->startElement('c:showVal');
				$objWriter->writeAttribute('val', '1');
				$objWriter->endElement();
				
			$objWriter->endElement();
			
			// c:gapWidth
			$objWriter->startElement('c:gapWidth');
			$objWriter->writeAttribute('val', '75');
			$objWriter->endElement();
			
			// c:shape
			$objWriter->startElement('c:shape');
			$objWriter->writeAttribute('val', 'box');
			$objWriter->endElement();
			
			// c:axId
			$objWriter->startElement('c:axId');
			$objWriter->writeAttribute('val', '52743552');
			$objWriter->endElement();
			
			// c:axId
			$objWriter->startElement('c:axId');
			$objWriter->writeAttribute('val', '52749440');
			$objWriter->endElement();
			
			// c:axId
			$objWriter->startElement('c:axId');
			$objWriter->writeAttribute('val', '0');
			$objWriter->endElement();
		
		$objWriter->endElement();
	}
	
	/**
	 * Write Type Pie3D
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 			$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Shape_Chart_Type_Pie3D	$subject
	 * @throws 	Exception
	 */
	protected function _writeTypePie3D(PHPPowerPoint_Shared_XMLWriter $objWriter, PHPPowerPoint_Shape_Chart_Type_Pie3D $subject) {
		// c:pie3DChart
		$objWriter->startElement('c:pie3DChart');
		
			// c:varyColors
			$objWriter->startElement('c:varyColors');
			$objWriter->writeAttribute('val', '1');
			$objWriter->endElement();
			
			// Write series
			$seriesIndex = 0;
			foreach ($subject->getData() as $series) {
				// c:ser
				$objWriter->startElement('c:ser');
				
					// c:idx
					$objWriter->startElement('c:idx');
					$objWriter->writeAttribute('val', $seriesIndex);
					$objWriter->endElement();
					
					// c:order
					$objWriter->startElement('c:order');
					$objWriter->writeAttribute('val', $seriesIndex);
					$objWriter->endElement();
					
					// c:tx
					$objWriter->startElement('c:tx');
					$objWriter->writeElement('c:v', $series->getTitle());
					$objWriter->endElement();
					
					// c:explosion
					$objWriter->startElement('c:explosion');
					$objWriter->writeAttribute('val', '20');
					$objWriter->endElement();
					
					// c:dLbls
					$objWriter->startElement('c:dLbls');

						// c:showPercent
						$objWriter->startElement('c:showPercent');
						$objWriter->writeAttribute('val', '1');
						$objWriter->endElement();		
						
						// c:showLeaderLines
						$objWriter->startElement('c:showLeaderLines');
						$objWriter->writeAttribute('val', '1');
						$objWriter->endElement();		
					
					$objWriter->endElement();
					
					// Write X axis data
					$axisXData = array_keys($series->getValues());
					
					// c:cat
					$objWriter->startElement('c:cat');
					
						// c:strLit / c:numLit
						if (is_int($axisXData[0]) || is_float($axisXData[0])) {
							$objWriter->startElement('c:numLit');
						} else {
							$objWriter->startElement('c:strLit');
						}
						
							// c:ptCount
							$objWriter->startElement('c:ptCount');
							$objWriter->writeAttribute('val', count($axisXData));
							$objWriter->endElement();
							
							// Add points
							for ($i = 0; $i < count($axisXData); $i++) {
								// c:pt
								$objWriter->startElement('c:pt');
								$objWriter->writeAttribute('idx', $i);
								$objWriter->writeElement('c:v', $axisXData[$i]);
								$objWriter->endElement();
							}
			
						$objWriter->endElement();
					
					$objWriter->endElement();
					
					// Write Y axis data
					$axisYData = array_values($series->getValues());
					
					// c:val
					$objWriter->startElement('c:val');
					
						// c:strLit / c:numLit
						if (is_int($axisYData[0]) || is_float($axisYData[0])) {
							$objWriter->startElement('c:numLit');
						} else {
							$objWriter->startElement('c:strLit');
						}
						
							// c:ptCount
							$objWriter->startElement('c:ptCount');
							$objWriter->writeAttribute('val', count($axisYData));
							$objWriter->endElement();
							
							// Add points
							for ($i = 0; $i < count($axisYData); $i++) {
								// c:pt
								$objWriter->startElement('c:pt');
								$objWriter->writeAttribute('idx', $i);
								$objWriter->writeElement('c:v', $axisYData[$i]);
								$objWriter->endElement();
							}
			
						$objWriter->endElement();
					
					$objWriter->endElement();
				
				$objWriter->endElement();
				
				++$seriesIndex;
			}

			// c:dLbls
			$objWriter->startElement('c:dLbls');
			
				// c:showVal
				$objWriter->startElement('c:showPercent');
				$objWriter->writeAttribute('val', '1');
				$objWriter->endElement();
				
			$objWriter->endElement();
		
		$objWriter->endElement();
	}
}