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
 * PHPPowerPoint_Writer_PowerPoint2007_Rels
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_Rels extends PHPPowerPoint_Writer_PowerPoint2007_WriterPart
{
	/**
	 * Write relationships to XML format
	 *
	 * @param 	PHPPowerPoint	$pPHPPowerPoint
	 * @return 	string 		XML Output
	 * @throws 	Exception
	 */
	public function writeRelationships(PHPPowerPoint $pPHPPowerPoint = null)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// Relationships
		$objWriter->startElement('Relationships');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

			// Relationship docProps/app.xml
			$this->_writeRelationship(
				$objWriter,
				3,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
				'docProps/app.xml'
			);

			// Relationship docProps/core.xml
			$this->_writeRelationship(
				$objWriter,
				2,
				'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
				'docProps/core.xml'
			);

			// Relationship ppt/presentation.xml
			$this->_writeRelationship(
				$objWriter,
				1,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
				'ppt/presentation.xml'
			);

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write presentation relationships to XML format
	 *
	 * @param 	PHPPowerPoint	$pPHPPowerPoint
	 * @return 	string 		XML Output
	 * @throws 	Exception
	 */
	public function writePresentationRelationships(PHPPowerPoint $pPHPPowerPoint = null)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// Relationships
		$objWriter->startElement('Relationships');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

			// Relation id
			$relationId = 1;
			
			// Add slide masters
			$masterSlides = $this->getParentWriter()->getLayoutPack()->getMasterSlides();
			foreach ($masterSlides as $masterSlide) {
				// Relationship slideMasters/slideMasterX.xml
				$this->_writeRelationship(
					$objWriter,
					$relationId++,
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
					'slideMasters/slideMaster' . $masterSlide['masterid'] . '.xml'
				);
			}
			
			// Add slide theme (only one!)
			// Relationship theme/theme1.xml
			$this->_writeRelationship(
				$objWriter,
				$relationId++,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
				'theme/theme1.xml'
			);

			// Relationships with slides
			$slideCount = $pPHPPowerPoint->getSlideCount();
			for ($i = 0; $i < $slideCount; ++$i) {
				$this->_writeRelationship(
					$objWriter,
					($i + $relationId),
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
					'slides/slide' . ($i + 1) . '.xml'
				);
			}

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write slide master relationships to XML format
	 *
	 * @param   int				$masterId			Master slide id
	 * @return 	string 			XML Output
	 * @throws 	Exception
	 */
	public function writeSlideMasterRelationships($masterId = 1)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// Relationships
		$objWriter->startElement('Relationships');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

			// Keep content id
			$contentId = 0;
			
			// Lookup layouts
			$layoutPack = $this->getParentWriter()->getLayoutPack();
			$layouts = array();
			foreach ($layoutPack->getLayouts() as $key => $layout) {
				if ($layout['masterid'] == $masterId) {
					$layouts[$key] = $layout;
				}
			}
			
			// Write slideLayout relationships
			foreach ($layouts as $key => $layout) {
				$this->_writeRelationship(
					$objWriter,
					++$contentId,
					'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
					'../slideLayouts/slideLayout' . $key . '.xml'
				);
			}

			// Relationship theme/theme1.xml
			$this->_writeRelationship(
				$objWriter,
				++$contentId,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
				'../theme/theme' . $masterId . '.xml'
			);

			// Other relationships
			$otherRelations = $layoutPack->getMasterSlideRelations();
			foreach ($otherRelations as $otherRelation)
			{
				if ($otherRelation['masterid'] == $masterId) {
					$this->_writeRelationship(
						$objWriter,
						++$contentId,
						$otherRelation['type'],
						$otherRelation['target']
					);
				}
			}

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write slide layout relationships to XML format
	 *
	 * @param	int				$masterId
	 * @return 	string 			XML Output
	 * @throws 	Exception
	 */
	public function writeSlideLayoutRelationships($slideLayoutIndex, $masterId = 1)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// Layout pack
		$layoutPack		= $this->getParentWriter()->getLayoutPack();

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// Relationships
		$objWriter->startElement('Relationships');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

			// Write slideMaster relationship
			$this->_writeRelationship(
				$objWriter,
				1,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
				'../slideMasters/slideMaster' . $masterId . '.xml'
			);

			// Other relationships
			$otherRelations = $layoutPack->getLayoutRelations();
			foreach ($otherRelations as $otherRelation)
			{
				if ($otherRelation['layoutId'] == $slideLayoutIndex)
				{
					$this->_writeRelationship(
						$objWriter,
						$otherRelation['id'],
						$otherRelation['type'],
						$otherRelation['target']
					);
				}
			}

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write theme relationships to XML format
	 *
	 * @param   int				$masterId			Master slide id
	 * @return 	string 			XML Output
	 * @throws 	Exception
	 */
	public function writeThemeRelationships($masterId = 1)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// Layout pack
		$layoutPack		= $this->getParentWriter()->getLayoutPack();

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// Relationships
		$objWriter->startElement('Relationships');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

    		// Other relationships
    		$otherRelations = $layoutPack->getThemeRelations();
    		foreach ($otherRelations as $otherRelation)
    		{
    			if ($otherRelation['masterid'] == $masterId) {
	    			$this->_writeRelationship(
	    				$objWriter,
	    				$otherRelation['id'],
	    				$otherRelation['type'],
	    				$otherRelation['target']
	    			);
    			}
    		}

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write slide relationships to XML format
	 *
	 * @param 	PHPPowerPoint_Slide		$pSlide
	 * @param 	int						$pSlideId
	 * @return 	string 					XML Output
	 * @throws 	Exception
	 */
	public function writeSlideRelationships(PHPPowerPoint_Slide $pSlide = null, $pSlideId = 1)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
		}

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// Relationships
		$objWriter->startElement('Relationships');
		$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

			// Starting relation id
			$relId = 1;

			// Write slideLayout relationship
			$layoutPack		= $this->getParentWriter()->getLayoutPack();
			$layoutIndex	= $layoutPack->findlayoutIndex( $pSlide->getSlideLayout(), $pSlide->getSlideMasterId() );

			$this->_writeRelationship(
				$objWriter,
				$relId++,
				'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
				'../slideLayouts/slideLayout' . ($layoutIndex + 1) . '.xml'
			);

			// Write drawing relationships?
			if ($pSlide->getShapeCollection()->count() > 0) {
				// Loop trough images and write relationships
				$iterator = $pSlide->getShapeCollection()->getIterator();
				while ($iterator->valid()) {
					if ($iterator->current() instanceof PHPPowerPoint_Shape_Drawing
						|| $iterator->current() instanceof PHPPowerPoint_Shape_MemoryDrawing) {
						// Write relationship for image drawing
						$this->_writeRelationship(
							$objWriter,
							$relId,
							'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
							'../media/' . str_replace(' ', '', $iterator->current()->getIndexedFilename())
						);

						$iterator->current()->__relationId = 'rId' . $relId;
						
						++$relId;
					} else if ($iterator->current() instanceof PHPPowerPoint_Shape_Chart) {
						// Write relationship for chart drawing
						$this->_writeRelationship(
							$objWriter,
							$relId,
							'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
							'../charts/' . $iterator->current()->getIndexedFilename()
						);

						$iterator->current()->__relationId = 'rId' . $relId;
						
						++$relId;
					}

					$iterator->next();
				}
			}
			
			// Write hyperlink relationships?
			if ($pSlide->getShapeCollection()->count() > 0) {
				// Loop trough hyperlinks and write relationships
				$iterator = $pSlide->getShapeCollection()->getIterator();
				while ($iterator->valid()) {
					// Hyperlink on shape
					if ($iterator->current()->hasHyperlink()) {
						// Write relationship for hyperlink
						$hyperlink = $iterator->current()->getHyperlink();
						$hyperlink->__relationId = 'rId' . $relId;
						
						if (!$hyperlink->isInternal()) {
							$this->_writeRelationship(
								$objWriter,
								$relId,
								'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
								$hyperlink->getUrl(),
								'External'
							);
						} else {
							$this->_writeRelationship(
								$objWriter,
								$relId,
								'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
								'slide' . $hyperlink->getSlideNumber() . '.xml'
							);
						}

						++$relId;
					}
					
					// Hyperlink on rich text run
					if ($iterator->current() instanceof PHPPowerPoint_Shape_RichText) {
						foreach ($iterator->current()->getParagraphs() as $paragraph) {
							foreach ($paragraph->getRichTextElements() as $element) {
								if ($element instanceof PHPPowerPoint_Shape_RichText_Run
		           					|| $element instanceof PHPPowerPoint_Shape_RichText_TextElement)
				           		{
				           			if ($element->hasHyperlink()) {
										// Write relationship for hyperlink
										$hyperlink = $element->getHyperlink();
										$hyperlink->__relationId = 'rId' . $relId;
										
										if (!$hyperlink->isInternal()) {
											$this->_writeRelationship(
												$objWriter,
												$relId,
												'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
												$hyperlink->getUrl(),
												'External'
											);
										} else {
											$this->_writeRelationship(
												$objWriter,
												$relId,
												'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
												'slide' . $hyperlink->getSlideNumber() . '.xml'
											);
										}
				
										++$relId;
									}
								}
							}
						}
	       			}

					$iterator->next();
				}
			}

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write Override content type
	 *
	 * @param 	PHPPowerPoint_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	int							$pId			Relationship ID. rId will be prepended!
	 * @param 	string						$pType			Relationship type
	 * @param 	string 						$pTarget		Relationship target
	 * @param 	string 						$pTargetMode	Relationship target mode
	 * @throws 	Exception
	 */
	private function _writeRelationship(PHPPowerPoint_Shared_XMLWriter $objWriter = null, $pId = 1, $pType = '', $pTarget = '', $pTargetMode = '')
	{
		if ($pType != '' && $pTarget != '') {
			if (strpos($pId, 'rId') === false) {
				$pId = 'rId' . $pId;
			}

			// Write relationship
			$objWriter->startElement('Relationship');
			$objWriter->writeAttribute('Id', 		$pId);
			$objWriter->writeAttribute('Type', 		$pType);
			$objWriter->writeAttribute('Target',	$pTarget);

			if ($pTargetMode != '') {
				$objWriter->writeAttribute('TargetMode',	$pTargetMode);
			}

			$objWriter->endElement();
		} else {
			throw new Exception("Invalid parameters passed.");
		}
	}
}
