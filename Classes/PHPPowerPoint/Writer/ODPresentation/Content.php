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
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Writer_ODPresentation_Content
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_ODPresentation
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_ODPresentation_Content extends PHPPowerPoint_Writer_ODPresentation_WriterPart
{
	/**
	 * Write content file to XML format
	 *
	 * @param 	PHPPowerPoint $pPHPPowerPoint
	 * @return 	string 						XML Output
	 * @throws 	Exception
	 */
	public function writeContent(PHPPowerPoint $pPHPPowerPoint = null)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8');

		// office:document-content
		$objWriter->startElement('office:document-content');
		$objWriter->writeAttribute('xmlns:office', 'urn:oasis:names:tc:opendocument:xmlns:office:1.0');
		$objWriter->writeAttribute('xmlns:style', 'urn:oasis:names:tc:opendocument:xmlns:style:1.0');
		$objWriter->writeAttribute('xmlns:text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
		$objWriter->writeAttribute('xmlns:table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
		$objWriter->writeAttribute('xmlns:draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
		$objWriter->writeAttribute('xmlns:fo', 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0');
		$objWriter->writeAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
		$objWriter->writeAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
		$objWriter->writeAttribute('xmlns:meta', 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0');
		$objWriter->writeAttribute('xmlns:number', 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0');
		$objWriter->writeAttribute('xmlns:presentation', 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0');
		$objWriter->writeAttribute('xmlns:svg', 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0');
		$objWriter->writeAttribute('xmlns:chart', 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0');
		$objWriter->writeAttribute('xmlns:dr3d', 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0');
		$objWriter->writeAttribute('xmlns:math', 'http://www.w3.org/1998/Math/MathML');
		$objWriter->writeAttribute('xmlns:form', 'urn:oasis:names:tc:opendocument:xmlns:form:1.0');
		$objWriter->writeAttribute('xmlns:script', 'urn:oasis:names:tc:opendocument:xmlns:script:1.0');
		$objWriter->writeAttribute('xmlns:ooo', 'http://openoffice.org/2004/office');
		$objWriter->writeAttribute('xmlns:ooow', 'http://openoffice.org/2004/writer');
		$objWriter->writeAttribute('xmlns:oooc', 'http://openoffice.org/2004/calc');
		$objWriter->writeAttribute('xmlns:dom', 'http://www.w3.org/2001/xml-events');
		$objWriter->writeAttribute('xmlns:xforms', 'http://www.w3.org/2002/xforms');
		$objWriter->writeAttribute('xmlns:xsd', 'http://www.w3.org/2001/XMLSchema');
		$objWriter->writeAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
		$objWriter->writeAttribute('xmlns:smil', 'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0');
		$objWriter->writeAttribute('xmlns:anim', 'urn:oasis:names:tc:opendocument:xmlns:animation:1.0');
		$objWriter->writeAttribute('xmlns:rpt', 'http://openoffice.org/2005/report');
		$objWriter->writeAttribute('xmlns:of', 'urn:oasis:names:tc:opendocument:xmlns:of:1.2');
		$objWriter->writeAttribute('xmlns:rdfa', 'http://docs.oasis-open.org/opendocument/meta/rdfa#');
		$objWriter->writeAttribute('xmlns:field', 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0');
		$objWriter->writeAttribute('office:version', '1.2');
			
			// office:automatic-styles
			$objWriter->startElement('office:automatic-styles');
			
			$slideCount = $pPHPPowerPoint->getSlideCount();
			for ($i = 0; $i < $slideCount; ++$i) {
				$pSlide = $pPHPPowerPoint->getSlide($i);
				// Images
				$shapes 	= 	$pSlide->getShapeCollection();
				$shapeId 	= 	0;
				foreach ($shapes as $shape)
				{
					// Increment $shapeId
					++$shapeId;
			
					// Check type
					if ($shape instanceof PHPPowerPoint_Shape_RichText) {
						
						// style:style
						$objWriter->startElement('style:style');
						$objWriter->writeAttribute('style:name', 'gr'.$shapeId);
						$objWriter->writeAttribute('style:family', 'graphic');
						$objWriter->writeAttribute('style:parent-style-name', 'standard');
							// style:graphic-properties
							$objWriter->startElement('style:graphic-properties');
							$objWriter->writeAttribute('draw:auto-grow-height', 'true');
							$objWriter->writeAttribute('fo:wrap-option', 'wrap');
							
							$objWriter->endElement();
						$objWriter->endElement();
						
						$paragraphs = $shape->getParagraphs();
						$paragraphId = 0;
						foreach ($paragraphs as $paragraph){
							++$paragraphId;
							
							// style:style
							$objWriter->startElement('style:style');
							$objWriter->writeAttribute('style:name', 'P'.$shapeId.'_'.$paragraphId);
							$objWriter->writeAttribute('style:family', 'paragraph');
								// style:paragraph-properties
								$objWriter->startElement('style:paragraph-properties');
								switch ($paragraph->getAlignment()->getHorizontal()) {
									case PHPPowerPoint_Style_Alignment::HORIZONTAL_LEFT:		$objWriter->writeAttribute('fo:text-align', 'left');	break;
									case PHPPowerPoint_Style_Alignment::HORIZONTAL_RIGHT:		$objWriter->writeAttribute('fo:text-align', 'right');	break;
									case PHPPowerPoint_Style_Alignment::HORIZONTAL_CENTER:		$objWriter->writeAttribute('fo:text-align', 'center');	break;
									case PHPPowerPoint_Style_Alignment::HORIZONTAL_JUSTIFY:		$objWriter->writeAttribute('fo:text-align', 'justify');	break;
									case PHPPowerPoint_Style_Alignment::HORIZONTAL_DISTRIBUTED:	$objWriter->writeAttribute('fo:text-align', 'justify');	break;
									default:													$objWriter->writeAttribute('fo:text-align', 'left');	break;
								}
									
								$objWriter->endElement();
							$objWriter->endElement();
							
							$richtexts = $paragraph->getRichTextElements();
							$richtextId = 0;
							foreach ($richtexts as $richtext) {
								++$richtextId;
								
								$elements = $paragraph->getRichTextElements();
								$elementId = 0;
								foreach ($elements as $element) {
									// Not a line break
									if ($element instanceof PHPPowerPoint_Shape_RichText_Run) {
										// style:style
										$objWriter->startElement('style:style');
										$objWriter->writeAttribute('style:name', 'T'.$shapeId.'_'.$paragraphId.'_'.$richtextId.'_'.$elementId);
										$objWriter->writeAttribute('style:family', 'text');
											// style:text-properties
											$objWriter->startElement('style:text-properties');
											$objWriter->writeAttribute('fo:color', '#'.$element->getFont()->getColor()->getRGB());
											$objWriter->writeAttribute('fo:font-family', $element->getFont()->getName());
											$objWriter->writeAttribute('fo:font-size', $element->getFont()->getSize().'pt');
											// @todo : fo:font-style
											if($element->getFont()->getBold()){
												$objWriter->writeAttribute('fo:font-weight', 'bold');
											}
											// @todo : style:text-underline-style
											$objWriter->endElement();
										$objWriter->endElement();
									}
								}
							}
						}
					}
					if ($shape instanceof PHPPowerPoint_Shape_BaseDrawing) {
						if ($shape->getShadow()->getVisible()) {
							// style:style
							$objWriter->startElement('style:style');
							$objWriter->writeAttribute('style:name', 'gr'.$shapeId);
							$objWriter->writeAttribute('style:family', 'graphic');
							$objWriter->writeAttribute('style:parent-style-name', 'standard');
							
								// style:graphic-properties
								$objWriter->startElement('style:graphic-properties');
								$objWriter->writeAttribute('draw:stroke', 'none');
								$objWriter->writeAttribute('draw:fill', 'none');
								$objWriter->writeAttribute('draw:shadow', 'visible');
								$objWriter->writeAttribute('draw:shadow-color', '#'.$shape->getShadow()->getColor()->getRGB());
								if($shape->getShadow()->getDirection() == 0 || $shape->getShadow()->getDirection() == 360) {
									$objWriter->writeAttribute('draw:shadow-offset-x', 	   PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
									$objWriter->writeAttribute('draw:shadow-offset-y', '0cm');
								}
								elseif($shape->getShadow()->getDirection() == 45) {
									$objWriter->writeAttribute('draw:shadow-offset-x', 	   PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
									$objWriter->writeAttribute('draw:shadow-offset-y', 	   PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
								}
								elseif($shape->getShadow()->getDirection() == 90) {
									$objWriter->writeAttribute('draw:shadow-offset-x', '0cm');
									$objWriter->writeAttribute('draw:shadow-offset-y', 	   PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
								}
								elseif($shape->getShadow()->getDirection() == 135) {
									$objWriter->writeAttribute('draw:shadow-offset-x', '-'.PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
									$objWriter->writeAttribute('draw:shadow-offset-y', 	   PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
								}
								elseif($shape->getShadow()->getDirection() == 180) {
									$objWriter->writeAttribute('draw:shadow-offset-x', '-'.PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
									$objWriter->writeAttribute('draw:shadow-offset-y', '0cm');
								}
								elseif($shape->getShadow()->getDirection() == 225) {
									$objWriter->writeAttribute('draw:shadow-offset-x', '-'.PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
									$objWriter->writeAttribute('draw:shadow-offset-y', '-'.PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
								}
								elseif($shape->getShadow()->getDirection() == 270) {
									$objWriter->writeAttribute('draw:shadow-offset-x', '0cm');
									$objWriter->writeAttribute('draw:shadow-offset-y', '-'.PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
								}
								elseif($shape->getShadow()->getDirection() == 315) {
									$objWriter->writeAttribute('draw:shadow-offset-x', 	   PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
									$objWriter->writeAttribute('draw:shadow-offset-y', '-'.PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getShadow()->getDistance()).'cm');
								}
								
								$objWriter->writeAttribute('draw:shadow-opacity', (100 - $shape->getShadow()->getAlpha()).'%');
								$objWriter->writeAttribute('style:mirror', 'none');
								$objWriter->endElement();
							
							$objWriter->endElement();
						}
					}
				}
			}
			
			$objWriter->endElement();
			
			// office:body
			$objWriter->startElement('office:body');
				// office:presentation
				$objWriter->startElement('office:presentation');
			
				// Write slides
				$slideCount = $pPHPPowerPoint->getSlideCount();
				for ($i = 0; $i < $slideCount; ++$i) {
					$pSlide = $pPHPPowerPoint->getSlide($i);
					$objWriter->startElement('draw:page');
					$objWriter->writeAttribute('draw:name', 'page'.$i);
					$objWriter->writeAttribute('draw:master-page-name', 'Default');
					
						// Images
						$shapes 	= 	$pSlide->getShapeCollection();
						$shapeId 	= 	0;
						foreach ($shapes as $shape)
						{
							// Increment $shapeId
							++$shapeId;
						
							// Check type
							if ($shape instanceof PHPPowerPoint_Shape_RichText)
							{
								$this->_writeTxt($objWriter, $shape, $shapeId);
							}
							/*else if ($shape instanceof PHPPowerPoint_Shape_Table)
							{
								$this->_writeTable($objWriter, $shape, $shapeId);
							}
							else if ($shape instanceof PHPPowerPoint_Shape_Line)
							{
								$this->_writeLineShape($objWriter, $shape, $shapeId);
							}
							else if ($shape instanceof PHPPowerPoint_Shape_Chart)
							{
								$this->_writeChart($objWriter, $shape, $shapeId);
							}
							else*/ if ($shape instanceof PHPPowerPoint_Shape_BaseDrawing)
							{
								$this->_writePic($objWriter, $shape, $shapeId);
							}
						}
						
						
						$objWriter->endElement();
					$objWriter->endElement();
				}
				
				/* 	<draw:page draw:name="page1" draw:style-name="dp1" draw:master-page-name="Default">
						<presentation:notes draw:style-name="dp2"><draw:page-thumbnail draw:style-name="gr1" draw:layer="layout" svg:width="14.848cm" svg:height="11.136cm" svg:x="3.075cm" svg:y="2.257cm" draw:page-number="1" presentation:class="page"/>
							<draw:frame presentation:style-name="pr1" draw:layer="layout" svg:width="16.799cm" svg:height="13.365cm" svg:x="2.1cm" svg:y="14.107cm" presentation:class="notes" presentation:placeholder="true">
								<draw:text-box/>
							</draw:frame>
						</presentation:notes>
					</draw:page>*/
				$objWriter->endElement();
			$objWriter->endElement();
		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}
	
	function _writePic(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Shape_BaseDrawing $shape = null, $shapeId){
		$objWriter->startElement('draw:frame');
		$objWriter->writeAttribute('draw:name', $shape->getName());
		$objWriter->writeAttribute('svg:width', number_format(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getWidth()),3).'cm');
		$objWriter->writeAttribute('svg:height', number_format(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getHeight()),3).'cm');
		$objWriter->writeAttribute('svg:x', number_format(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getOffsetX()),3).'cm');
		$objWriter->writeAttribute('svg:y', number_format(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getOffsetY()),3).'cm');
		if ($shape->getShadow()->getVisible()) {
			$objWriter->writeAttribute('draw:style-name', 'gr'.$shapeId);
		}

			$objWriter->startElement('draw:image');
			$objWriter->writeAttribute('xlink:href', 'Pictures/'.str_replace(' ', '_', $shape->getIndexedFilename()));
			$objWriter->writeAttribute('xlink:type', 'simple');
			$objWriter->writeAttribute('xlink:show', 'embed');
			$objWriter->writeAttribute('xlink:actuate', 'onLoad');
			
			$objWriter->writeElement('text:p');
		
			$objWriter->endElement();

		$objWriter->endElement();
	}

	function _writeTxt(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Shape_RichText $shape = null, $shapeId){
		// draw:custom-shape
		$objWriter->startElement('draw:custom-shape');
		$objWriter->writeAttribute('draw:style-name', 'gr'.$shapeId);
		$objWriter->writeAttribute('svg:width', number_format(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getWidth()),3).'cm');
		$objWriter->writeAttribute('svg:height', number_format(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getHeight()),3).'cm');
		$objWriter->writeAttribute('svg:x', number_format(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getOffsetX()),3).'cm');
		$objWriter->writeAttribute('svg:y', number_format(PHPPowerPoint_Shared_Drawing::pixelsToCentimeters($shape->getOffsetY()),3).'cm');
		
			$paragraphs = $shape->getParagraphs();
			$paragraphId = 0;
			foreach ($paragraphs as $paragraph){
				++$paragraphId;
				// text:p
				$objWriter->startElement('text:p');
				$objWriter->writeAttribute('text:style-name', 'P'.$shapeId.'_'.$paragraphId);
					// Loop trough rich text elements
					$richtexts = $paragraph->getRichTextElements();
					$richtextId = 0;
					foreach ($richtexts as $richtext) {
						
						++$richtextId;
						
						$elements = $paragraph->getRichTextElements();
						$elementId = 0;
						foreach ($elements as $element) {
							// Not a line break
							if ($element instanceof PHPPowerPoint_Shape_RichText_TextElement || $element instanceof PHPPowerPoint_Shape_RichText_Run) {
								// text:span
								$objWriter->startElement('text:span');
								if($element instanceof PHPPowerPoint_Shape_RichText_Run){
									$objWriter->writeAttribute('text:style-name', 'T'.$shapeId.'_'.$paragraphId.'_'.$richtextId.'_'.$elementId);
								}
								$objWriter->text($richtext->getText());
								$objWriter->endElement();
							}
						}
					}
				$objWriter->endElement();
			}
		$objWriter->endElement();
	}
}
