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
 * @package	PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	##VERSION##, ##DATE##
 */


/**
 * PHPPowerPoint_Writer_PowerPoint2007_Slide
 *
 * @category   PHPPowerPoint
 * @package	PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2006 - 2009 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_Slide extends PHPPowerPoint_Writer_PowerPoint2007_WriterPart
{
	/**
	 * Write slide to XML format
	 *
	 * @param	PHPPowerPoint_Slide		$pSlide
	 * @return	string					XML Output
	 * @throws	Exception
	 */
	public function writeSlide(PHPPowerPoint_Slide $pSlide = null)
	{
		// Check slide
		if (is_null($pSlide))
			throw new Exception("Invalid PHPPowerPoint_Slide object passed.");

		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// p:sld
		$objWriter->startElement('p:sld');
		$objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
		$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		$objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

    		// p:cSld
    		$objWriter->startElement('p:cSld');

    			// p:spTree
    			$objWriter->startElement('p:spTree');

    				// p:nvGrpSpPr
    				$objWriter->startElement('p:nvGrpSpPr');

        				// p:cNvPr
        				$objWriter->startElement('p:cNvPr');
        				$objWriter->writeAttribute('id', '1');
        				$objWriter->writeAttribute('name', '');
        				$objWriter->endElement();

        				// p:cNvGrpSpPr
        				$objWriter->writeElement('p:cNvGrpSpPr', null);

        				// p:nvPr
        				$objWriter->writeElement('p:nvPr', null);

    				$objWriter->endElement();

    				// p:grpSpPr
    				$objWriter->startElement('p:grpSpPr');

        				// a:xfrm
        				$objWriter->startElement('a:xfrm');

            				// a:off
            				$objWriter->startElement('a:off');
            				$objWriter->writeAttribute('x', '0');
            				$objWriter->writeAttribute('y', '0');
            				$objWriter->endElement();

            				// a:ext
            				$objWriter->startElement('a:ext');
            				$objWriter->writeAttribute('cx', '0');
            				$objWriter->writeAttribute('cy', '0');
            				$objWriter->endElement();

            				// a:chOff
            				$objWriter->startElement('a:chOff');
            				$objWriter->writeAttribute('x', '0');
            				$objWriter->writeAttribute('y', '0');
            				$objWriter->endElement();

            				// a:chExt
            				$objWriter->startElement('a:chExt');
            				$objWriter->writeAttribute('cx', '0');
            				$objWriter->writeAttribute('cy', '0');
            				$objWriter->endElement();

        				$objWriter->endElement();

        			$objWriter->endElement();

        			// Loop shapes
        			$shapeId 	= 0;
        			$relationId = 1;
        			$shapes 	= $pSlide->getShapeCollection();
        			foreach ($shapes as $shape)
        			{
        				// Increment $shapeId
        				++$shapeId;

        				// Check type
        				if ($shape instanceof PHPPowerPoint_Shape_RichText)
        				{
        					$this->_writeTxt($objWriter, $shape, $shapeId);
        				}
        				else if ($shape instanceof PHPPowerPoint_Shape_Table)
        				{
        					$this->_writeTable($objWriter, $shape, $shapeId);
        				}
        				else if ($shape instanceof PHPPowerPoint_Shape_BaseDrawing)
        				{
        					// Picture --> $relationId
        					++$relationId;

        					$this->_writePic($objWriter, $shape, $shapeId, $relationId);
        				}
        			}
        			
        			// TODO

    			$objWriter->endElement();

    		$objWriter->endElement();

    		// p:clrMapOvr
    		$objWriter->startElement('p:clrMapOvr');

    			// a:masterClrMapping
    			$objWriter->writeElement('a:masterClrMapping', '');

    		$objWriter->endElement();

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write pic
	 *
	 * @param	PHPPowerPoint_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPPowerPoint_Shape_BaseDrawing		$shape
	 * @param	int									$shapeId
	 * @param	int									$relationId
	 * @throws	Exception
	 */
	private function _writePic(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Shape_BaseDrawing $shape = null, $shapeId, $relationId)
	{
		// p:pic
		$objWriter->startElement('p:pic');

			// p:nvPicPr
			$objWriter->startElement('p:nvPicPr');

				// p:cNvPr
				$objWriter->startElement('p:cNvPr');
                $objWriter->writeAttribute('id', $shapeId);
                $objWriter->writeAttribute('name', $shape->getName());
				$objWriter->writeAttribute('descr', $shape->getDescription());
                $objWriter->endElement();

                // p:cNvPicPr
                $objWriter->startElement('p:cNvPicPr');

                	// a:picLocks
                	$objWriter->startElement('a:picLocks');
                	$objWriter->writeAttribute('noChangeAspect', '1');
                	$objWriter->endElement();

                $objWriter->endElement();

                // p:nvPr
        		$objWriter->writeElement('p:nvPr', null);

			$objWriter->endElement();

			// p:blipFill
			$objWriter->startElement('p:blipFill');

				// a:blip
				$objWriter->startElement('a:blip');
				$objWriter->writeAttribute('r:embed', 'rId' . $relationId);
				$objWriter->endElement();

				// a:stretch
				$objWriter->startElement('a:stretch');
					$objWriter->writeElement('a:fillRect', null);
                $objWriter->endElement();

			$objWriter->endElement();

			// p:spPr
			$objWriter->startElement('p:spPr');

				// a:xfrm
				$objWriter->startElement('a:xfrm');
				$objWriter->writeAttribute('rot', PHPPowerPoint_Shared_Drawing::degreesToAngle($shape->getRotation()));

					// a:off
					$objWriter->startElement('a:off');
					$objWriter->writeAttribute('x', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getOffsetX()));
                    $objWriter->writeAttribute('y', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getOffsetY()));
                    $objWriter->endElement();

                    // a:ext
                    $objWriter->startElement('a:ext');
                    $objWriter->writeAttribute('cx', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getWidth()));
                    $objWriter->writeAttribute('cy', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getHeight()));
                    $objWriter->endElement();

				$objWriter->endElement();

				// a:prstGeom
				$objWriter->startElement('a:prstGeom');
				$objWriter->writeAttribute('prst', 'rect');

					// a:avLst
					$objWriter->writeElement('a:avLst', null);

				$objWriter->endElement();

				if ($shape->getShadow()->getVisible()) {
					// a:effectLst
					$objWriter->startElement('a:effectLst');

						// a:outerShdw
						$objWriter->startElement('a:outerShdw');
						$objWriter->writeAttribute('blurRad', 		PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getShadow()->getBlurRadius()));
						$objWriter->writeAttribute('dist',			PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getShadow()->getDistance()));
						$objWriter->writeAttribute('dir',			PHPPowerPoint_Shared_Drawing::degreesToAngle($shape->getShadow()->getDirection()));
						$objWriter->writeAttribute('algn',			$shape->getShadow()->getAlignment());
						$objWriter->writeAttribute('rotWithShape', 	'0');

							// a:srgbClr
							$objWriter->startElement('a:srgbClr');
							$objWriter->writeAttribute('val',		$shape->getShadow()->getColor()->getRGB());

								// a:alpha
								$objWriter->startElement('a:alpha');
								$objWriter->writeAttribute('val', 	$shape->getShadow()->getAlpha() * 1000);
								$objWriter->endElement();

							$objWriter->endElement();

						$objWriter->endElement();

					$objWriter->endElement();
				}

            $objWriter->endElement();

        $objWriter->endElement();
	}

	/**
	 * Write txt
	 *
	 * @param	PHPPowerPoint_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPPowerPoint_Shape_RichText		$shape
	 * @param	int									$shapeId
	 * @throws	Exception
	 */
	private function _writeTxt(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Shape_RichText $shape = null, $shapeId)
	{
		// p:sp
		$objWriter->startElement('p:sp');

			// p:nvSpPr
			$objWriter->startElement('p:nvSpPr');

				// p:cNvPr
				$objWriter->startElement('p:cNvPr');
                $objWriter->writeAttribute('id', $shapeId);
                $objWriter->writeAttribute('name', '');
                $objWriter->endElement();

                // p:cNvSpPr
                $objWriter->startElement('p:cNvSpPr');
                $objWriter->writeAttribute('txBox', '1');
                $objWriter->endElement();

                // p:nvPr
        		$objWriter->writeElement('p:nvPr', null);

			$objWriter->endElement();

			// p:spPr
			$objWriter->startElement('p:spPr');

				// a:xfrm
				$objWriter->startElement('a:xfrm');
				$objWriter->writeAttribute('rot', PHPPowerPoint_Shared_Drawing::degreesToAngle($shape->getRotation()));

					// a:off
					$objWriter->startElement('a:off');
					$objWriter->writeAttribute('x', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getOffsetX()));
                    $objWriter->writeAttribute('y', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getOffsetY()));
                    $objWriter->endElement();

                    // a:ext
                    $objWriter->startElement('a:ext');
                    $objWriter->writeAttribute('cx', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getWidth()));
                    $objWriter->writeAttribute('cy', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getHeight()));
                    $objWriter->endElement();

				$objWriter->endElement();

				// a:prstGeom
				$objWriter->startElement('a:prstGeom');
				$objWriter->writeAttribute('prst', 'rect');
				$objWriter->endElement();

				if ($shape->getFill()->getFillType()==='solid') {
					// a:solidFill
					$objWriter->startElement('a:solidFill');
					$objWriter->startElement('a:srgbClr');
					$objWriter->writeAttribute('val', $shape->getFill()->getStartColor()->getRGB());
					$objWriter->endElement();
					$objWriter->endElement();
				} else {
					// a:noFill
					$objWriter->writeElement('a:noFill', null);
				}

				if ($shape->getShadow()->getVisible()) {
					// a:effectLst
					$objWriter->startElement('a:effectLst');

						// a:outerShdw
						$objWriter->startElement('a:outerShdw');
						$objWriter->writeAttribute('blurRad', 		PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getShadow()->getBlurRadius()));
						$objWriter->writeAttribute('dist',			PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getShadow()->getDistance()));
						$objWriter->writeAttribute('dir',			PHPPowerPoint_Shared_Drawing::degreesToAngle($shape->getShadow()->getDirection()));
						$objWriter->writeAttribute('algn',			$shape->getShadow()->getAlignment());
						$objWriter->writeAttribute('rotWithShape', 	'0');

							// a:srgbClr
							$objWriter->startElement('a:srgbClr');
							$objWriter->writeAttribute('val',		$shape->getShadow()->getColor()->getRGB());

								// a:alpha
								$objWriter->startElement('a:alpha');
								$objWriter->writeAttribute('val', 	$shape->getShadow()->getAlpha() * 1000);
								$objWriter->endElement();

							$objWriter->endElement();

						$objWriter->endElement();

					$objWriter->endElement();
				}

            $objWriter->endElement();

			// p:txBody
			$objWriter->startElement('p:txBody');

				// a:bodyPr
				$objWriter->startElement('a:bodyPr');
                $objWriter->writeAttribute('wrap', 'square');
                $objWriter->writeAttribute('rtlCol', '0');

                    // a:spAutoFit
            		$objWriter->writeElement('a:spAutoFit', null);

                $objWriter->endElement();

                // a:lstStyle
            	$objWriter->writeElement('a:lstStyle', null);

        		// Write paragraphs
        		$this->_writeParagraphs($objWriter, $shape->getParagraphs());

			$objWriter->endElement();

        $objWriter->endElement();
	}

	/**
	 * Write table
	 *
	 * @param	PHPPowerPoint_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPPowerPoint_Shape_Table    		$shape
	 * @param	int									$shapeId
	 * @throws	Exception
	 */
	private function _writeTable(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Shape_Table $shape = null, $shapeId)
	{
		// p:graphicFrame
		$objWriter->startElement('p:graphicFrame');

			// p:nvGraphicFramePr
			$objWriter->startElement('p:nvGraphicFramePr');

				// p:cNvPr
				$objWriter->startElement('p:cNvPr');
                $objWriter->writeAttribute('id', $shapeId);
                $objWriter->writeAttribute('name', $shape->getName());
				$objWriter->writeAttribute('descr', $shape->getDescription());
                $objWriter->endElement();

                // p:cNvGraphicFramePr
                $objWriter->startElement('p:cNvGraphicFramePr');

                	// a:graphicFrameLocks
                	$objWriter->startElement('a:graphicFrameLocks');
                	$objWriter->writeAttribute('noGrp', '1');
                	$objWriter->endElement();

                $objWriter->endElement();

                // p:nvPr
        		$objWriter->writeElement('p:nvPr', null);

			$objWriter->endElement();

			// p:xfrm
			$objWriter->startElement('p:xfrm');

				// a:off
				$objWriter->startElement('a:off');
				$objWriter->writeAttribute('x', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getOffsetX()));
                $objWriter->writeAttribute('y', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getOffsetY()));
                $objWriter->endElement();

                // a:ext
                $objWriter->startElement('a:ext');
                $objWriter->writeAttribute('cx', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getWidth()));
                $objWriter->writeAttribute('cy', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getHeight()));
                $objWriter->endElement();

			$objWriter->endElement();

			// a:graphic
			$objWriter->startElement('a:graphic');

				// a:graphicData
				$objWriter->startElement('a:graphicData');
				$objWriter->writeAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/table');

					// a:tbl
					$objWriter->startElement('a:tbl');

		                // a:tblPr
		                $objWriter->startElement('a:tblPr');
		                $objWriter->writeAttribute('firstRow', '1');
		                $objWriter->writeAttribute('bandRow', '1');
		                $objWriter->endElement();

						// a:tblGrid
						$objWriter->startElement('a:tblGrid');

							// Write cell widths
							for ($cell = 0; $cell < count($shape->getRow(0)->getCells()); $cell++)
							{
				                // a:gridCol
				                $objWriter->startElement('a:gridCol');
				                
				                // Calculate column width
				                $width = $shape->getRow(0)->getCell($cell)->getWidth();
				                if ($width == 0) {
				                	$colCount = count($shape->getRow(0)->getCells());
				                	$totalWidth = $shape->getWidth();
				                	$width = $totalWidth / $colCount;
				                }
				                
				                $objWriter->writeAttribute('w', PHPPowerPoint_Shared_Drawing::pixelsToEMU($width));
				                $objWriter->endElement();
							}

		                $objWriter->endElement();

		                // Colspan / rowspan containers
		                $colSpan = array();
		                $rowSpan = array();

		                // Write rows
		                for ($row = 0; $row < count($shape->getRows()); $row++)
		                {
							// a:tr
							$objWriter->startElement('a:tr');
							$objWriter->writeAttribute('h', PHPPowerPoint_Shared_Drawing::pixelsToEMU($shape->getRow($row)->getHeight()));

								// Write cells
								for ($cell = 0; $cell < count($shape->getRow($row)->getCells()); $cell++)
								{
									// Current cell
									$currentCell = $shape->getRow($row)->getCell($cell);

					                // a:tc
					                $objWriter->startElement('a:tc');
					                	// Colspan
										if ($currentCell->getColSpan() > 1)
						                {
						                	$objWriter->writeAttribute('gridSpan', $currentCell->getColSpan());
						                	$colSpan[$row] = $currentCell->getColSpan() - 1;
						                }
						                else if (isset($colSpan[$row]) && $colSpan[$row] > 0)
						                {
						                	$colSpan[$row]--;
						                	$objWriter->writeAttribute('hMerge', '1');
						                }

						                // Rowspan
										if ($currentCell->getRowSpan() > 1)
						                {
						                	$objWriter->writeAttribute('rowSpan', $currentCell->getRowSpan());
						                	$rowSpan[$cell] = $currentCell->getRowSpan() - 1;
						                }
										else if (isset($rowSpan[$cell]) && $rowSpan[$cell] > 0)
						                {
						                	$rowSpan[$cell]--;
						                	$objWriter->writeAttribute('vMerge', '1');
						                }


										// a:txBody
										$objWriter->startElement('a:txBody');

											// a:bodyPr
											$objWriter->startElement('a:bodyPr');
							                $objWriter->writeAttribute('wrap', 'square');
							                $objWriter->writeAttribute('rtlCol', '0');

							                    // a:spAutoFit
							            		$objWriter->writeElement('a:spAutoFit', null);

							                $objWriter->endElement();

							                // a:lstStyle
							            	$objWriter->writeElement('a:lstStyle', null);

							            	// Write paragraphs
							            	$this->_writeParagraphs($objWriter, $currentCell->getParagraphs());

										$objWriter->endElement();

						                // a:tcPr
						                $objWriter->startElement('a:tcPr');
						                	// Alignment (horizontal)
						                	$firstParagraph = $currentCell->getParagraph(0);
						                	$horizontalAlign = $firstParagraph->getAlignment()->getVertical();
						                	if ($horizontalAlign != PHPPowerPoint_Style_Alignment::VERTICAL_BASE && $horizontalAlign != PHPPowerPoint_Style_Alignment::VERTICAL_AUTO)
						                	{
						                		$objWriter->writeAttribute('anchor', $horizontalAlign);
						                	}

						                	// Borders
						                	$this->_writeBorders($objWriter, $currentCell->getBorders());

						                	// Fill
						                	$this->_writeFill($objWriter, $currentCell->getFill());

						                $objWriter->endElement();

					                $objWriter->endElement();
								}

			                $objWriter->endElement();
		                }

	                $objWriter->endElement();

                $objWriter->endElement();

			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write paragraphs
	 *
	 * @param	PHPPowerPoint_Shared_XMLWriter				$objWriter		XML Writer
	 * @param	PHPPowerPoint_Shape_RichText_Paragraph[]	$paragraphs
	 * @throws	Exception
	 */
	private function _writeParagraphs(PHPPowerPoint_Shared_XMLWriter $objWriter, $paragraphs)
	{
		// Loop trough paragraphs
        foreach ($paragraphs as $paragraph)
        {
			// a:p
			$objWriter->startElement('a:p');

		    	// a:pPr
	        	$objWriter->startElement('a:pPr');
	        	$objWriter->writeAttribute('algn', 		$paragraph->getAlignment()->getHorizontal());
	        	$objWriter->writeAttribute('fontAlgn', 	$paragraph->getAlignment()->getVertical());
	        	$objWriter->writeAttribute('marL', 		PHPPowerPoint_Shared_Drawing::pixelsToEMU($paragraph->getAlignment()->getMarginLeft()));
	        	$objWriter->writeAttribute('marR', 		PHPPowerPoint_Shared_Drawing::pixelsToEMU($paragraph->getAlignment()->getMarginRight()));
	        	$objWriter->writeAttribute('indent', 	PHPPowerPoint_Shared_Drawing::pixelsToEMU($paragraph->getAlignment()->getIndent()));
	        	$objWriter->writeAttribute('lvl', 		$paragraph->getAlignment()->getLevel());

	        	// Bullet type specified?
	        	if ($paragraph->getBulletStyle()->getBulletType() != PHPPowerPoint_Style_Bullet::TYPE_NONE)
	        	{
			   		// a:buFont
		    		$objWriter->startElement('a:buFont');
		    		$objWriter->writeAttribute('typeface', 		$paragraph->getBulletStyle()->getBulletFont());
		    		$objWriter->endElement();

		    		if ($paragraph->getBulletStyle()->getBulletType() == PHPPowerPoint_Style_Bullet::TYPE_BULLET)
		    		{
			      		// a:buChar
			   			$objWriter->startElement('a:buChar');
			   			$objWriter->writeAttribute('char', 			$paragraph->getBulletStyle()->getBulletChar());
			   			$objWriter->endElement();
		    		}
		    		else if ($paragraph->getBulletStyle()->getBulletType() == PHPPowerPoint_Style_Bullet::TYPE_NUMERIC)
		    		{
			     		// a:buAutoNum
			   			$objWriter->startElement('a:buAutoNum');
			   			$objWriter->writeAttribute('type', 			$paragraph->getBulletStyle()->getBulletNumericStyle());
			   			if ($paragraph->getBulletStyle()->getBulletNumericStartAt() != 1)
			   			{
			   				$objWriter->writeAttribute('startAt', 		$paragraph->getBulletStyle()->getBulletNumericStartAt());
			   			}
			   			$objWriter->endElement();
		    		}
	        	}

	       		$objWriter->endElement();

	      		// Loop trough rich text elements
	       		$elements = $paragraph->getRichTextElements();
	       		foreach ($elements as $element) {
	       			if ($element instanceof PHPPowerPoint_Shape_RichText_Break) {
	           			// a:br
	           			$objWriter->writeElement('a:br', null);
	       			}
	           		elseif ($element instanceof PHPPowerPoint_Shape_RichText_Run
	           				|| $element instanceof PHPPowerPoint_Shape_RichText_TextElement)
	           		{
	           			// a:r
	           			$objWriter->startElement('a:r');

	           				// a:rPr
	           				if ($element instanceof PHPPowerPoint_Shape_RichText_Run) {
	           					// a:rPr
	           					$objWriter->startElement('a:rPr');

	               					// Bold
	               					$objWriter->writeAttribute('b', ($element->getFont()->getBold() ? 'true' : 'false'));

	               					// Italic
	               					$objWriter->writeAttribute('i', ($element->getFont()->getItalic() ? 'true' : 'false'));

	               					// Strikethrough
	               					$objWriter->writeAttribute('strike', ($element->getFont()->getStrikethrough() ? 'sngStrike' : 'noStrike'));

	               					// Size
	               					$objWriter->writeAttribute('sz', ($element->getFont()->getSize() * 100));

	               					// Underline
	               					$objWriter->writeAttribute('u', $element->getFont()->getUnderline());

	               					// Superscript / subscript
	               					if ($element->getFont()->getSuperScript() || $element->getFont()->getSubScript()) {
	               						if ($element->getFont()->getSuperScript()) {
	               							$objWriter->writeAttribute('baseline', '30000');
	               						} else if ($element->getFont()->getSubScript()) {
	               							$objWriter->writeAttribute('baseline', '-25000');
	               						}
	               					}

	           						// Color - a:solidFill
	           						$objWriter->startElement('a:solidFill');

	           							// a:srgbClr
	               						$objWriter->startElement('a:srgbClr');
	               						$objWriter->writeAttribute('val', $element->getFont()->getColor()->getRGB());
	               						$objWriter->endElement();

	           						$objWriter->endElement();

	           						// Font - a:latin
	           						$objWriter->startElement('a:latin');
	           						$objWriter->writeAttribute('typeface', $element->getFont()->getName());
	           						$objWriter->endElement();

	           					$objWriter->endElement();
	          				}

	           				// t
	           				$objWriter->startElement('a:t');
	           				$objWriter->writeCData(PHPPowerPoint_Shared_String::ControlCharacterPHP2OOXML( $element->getText() ));
	           				$objWriter->endElement();

	           			$objWriter->endElement();
	           		}
	       		}

	    	$objWriter->endElement();
        }
	}

	/**
	 * Write Borders
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Style_Borders		$pBorders		Borders
	 * @throws 	Exception
	 */
	private function _writeBorders(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Style_Borders $pBorders = null)
	{
		$this->_writeBorder($objWriter, $pBorders->getLeft(), 			'L');
		$this->_writeBorder($objWriter, $pBorders->getRight(), 			'R');
		$this->_writeBorder($objWriter, $pBorders->getTop(), 			'T');
		$this->_writeBorder($objWriter, $pBorders->getBottom(), 		'B');
		$this->_writeBorder($objWriter, $pBorders->getDiagonalDown(), 	'TlToBr');
		$this->_writeBorder($objWriter, $pBorders->getDiagonalUp(), 	'BlToTr');
	}

	/**
	 * Write Border
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Style_Border		$pBorder		Border
	 * @param   string							$pElementName	Element name
	 * @throws 	Exception
	 */
	private function _writeBorder(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Style_Border $pBorder = null, $pElementName = 'L')
	{
		// Line style
		$lineStyle = $pBorder->getLineStyle();
		if ($lineStyle == PHPPowerPoint_Style_Border::LINE_NONE)
		{
			$lineStyle = PHPPowerPoint_Style_Border::LINE_SINGLE;
		}

		// Line width
		$lineWidth = 12700 * $pBorder->getLineWidth();

		// a:ln $pElementName
		$objWriter->startElement('a:ln' . $pElementName);
		$objWriter->writeAttribute('w', 	$lineWidth);
		$objWriter->writeAttribute('cap', 	'flat');
		$objWriter->writeAttribute('cmpd', 	$lineStyle);
		$objWriter->writeAttribute('algn', 	'ctr');

			// Fill?
			if ($pBorder->getLineStyle() == PHPPowerPoint_Style_Border::LINE_NONE)
			{
				// a:noFill
				$objWriter->writeElement('a:noFill', null);
			}
			else
			{
				// a:solidFill
				$objWriter->startElement('a:solidFill');

					// a:srgbClr
					$objWriter->startElement('a:srgbClr');
					$objWriter->writeAttribute('val', $pBorder->getColor()->getRGB());
					$objWriter->endElement();

				$objWriter->endElement();
			}

			// Dash
			$objWriter->startElement('a:prstDash');
			$objWriter->writeAttribute('val', $pBorder->getDashStyle());
			$objWriter->endElement();

			// a:round
			$objWriter->writeElement('a:round', null);

			// a:headEnd
			$objWriter->startElement('a:headEnd');
			$objWriter->writeAttribute('type', 	'none');
			$objWriter->writeAttribute('w', 	'med');
			$objWriter->writeAttribute('len', 	'med');
			$objWriter->endElement();

			// a:tailEnd
			$objWriter->startElement('a:tailEnd');
			$objWriter->writeAttribute('type', 	'none');
			$objWriter->writeAttribute('w', 	'med');
			$objWriter->writeAttribute('len', 	'med');
			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write Fill
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Style_Fill			$pFill			Fill style
	 * @throws 	Exception
	 */
	private function _writeFill(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Style_Fill $pFill = null)
	{
		// Is it a fill?
		if ($pFill->getFillType() == PHPPowerPoint_Style_Fill::FILL_NONE)
			return;

		// Check if this is a pattern type or gradient type
		if ($pFill->getFillType() == PHPPowerPoint_Style_Fill::FILL_GRADIENT_LINEAR
			|| $pFill->getFillType() == PHPPowerPoint_Style_Fill::FILL_GRADIENT_PATH) {
			// Gradient fill
			$this->_writeGradientFill($objWriter, $pFill);
		} else {
			// Pattern fill
			$this->_writePatternFill($objWriter, $pFill);
		}
	}

	/**
	 * Write Gradient Fill
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Style_Fill		$pFill			Fill style
	 * @throws 	Exception
	 */
	private function _writeGradientFill(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Style_Fill $pFill = null)
	{
		// a:gradFill
		$objWriter->startElement('a:gradFill');

			// a:gsLst
			$objWriter->startElement('a:gsLst');
				// a:gs
				$objWriter->startElement('a:gs');
				$objWriter->writeAttribute('pos', '0');

					// srgbClr
					$objWriter->startElement('a:srgbClr');
					$objWriter->writeAttribute('val', $pFill->getStartColor()->getRGB());
					$objWriter->endElement();

				$objWriter->endElement();

				// a:gs
				$objWriter->startElement('a:gs');
				$objWriter->writeAttribute('pos', '100000');

					// srgbClr
					$objWriter->startElement('a:srgbClr');
					$objWriter->writeAttribute('val', $pFill->getEndColor()->getRGB());
					$objWriter->endElement();

				$objWriter->endElement();

			$objWriter->endElement();

			// a:lin
			$objWriter->startElement('a:lin');
			$objWriter->writeAttribute('ang',    PHPPowerPoint_Shared_Drawing::degreesToAngle($pFill->getRotation()));
			$objWriter->writeAttribute('scaled', '0');
			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write Pattern Fill
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter			$objWriter 		XML Writer
	 * @param 	PHPPowerPoint_Style_Fill					$pFill			Fill style
	 * @throws 	Exception
	 */
	private function _writePatternFill(PHPPowerPoint_Shared_XMLWriter $objWriter = null, PHPPowerPoint_Style_Fill $pFill = null)
	{
		// a:pattFill
		$objWriter->startElement('a:pattFill');

			// fgClr
			$objWriter->startElement('a:fgClr');

				// srgbClr
				$objWriter->startElement('a:srgbClr');
				$objWriter->writeAttribute('val', $pFill->getStartColor()->getRGB());
				$objWriter->endElement();

			$objWriter->endElement();

			// bgClr
			$objWriter->startElement('a:bgClr');

				// srgbClr
				$objWriter->startElement('a:srgbClr');
				$objWriter->writeAttribute('val', $pFill->getEndColor()->getRGB());
				$objWriter->endElement();

			$objWriter->endElement();

		$objWriter->endElement();
	}
}
