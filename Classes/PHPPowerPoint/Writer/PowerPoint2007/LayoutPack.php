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
 * PHPPowerPoint_Writer_PowerPoint2007_LayoutPack
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
abstract class PHPPowerPoint_Writer_PowerPoint2007_LayoutPack
{
	/**
	 * Master slides
	 * 
	 * Structure:
	 * - masterid
	 * - body
	 * 
	 * @var array
	 */
	protected $_masterSlides = array();
	
	/**
	 * Master slide relations
	 * 
	 * Structure:
	 * - master id
	 * - id (relation id)
	 * - type
	 * - contentType
	 * - target (full path in OpenXML package)
	 * - contents (body)
	 *
	 * @var array
	 */
	protected $_masterSlideRelations = array();
	
	/**
	 * Themes
	 *
	 * Structure:
	 * - masterid
	 * - body
	 *
	 * @var array
	 */
	protected $_themes = '';
	
	/**
	 * Theme relations
	 * 
	 * Structure:
	 * - masterid
	 * - id (relation id)
	 * - type
	 * - contentType
	 * - target (full path in OpenXML package)
	 * - contents (body)
	 *
	 * @var array
	 */
	protected $_themeRelations = array();
	
	/**
	 * Array of slide layouts.
	 * 
	 * These are all an array consisting of:
	 * - masterid (int)
	 * - name (string)
	 * - body (string)
	 * 
	 * @var array
	 */
	protected $_layouts = array();
	
	/**
	 * Layout relations
	 * 
	 * Structure:
	 * - layoutId (referencing layout id in layouts array)
	 * - id (relation id)
	 * - type
	 * - contentType
	 * - target (full path in OpenXML package)
	 * - contents (body)
	 *
	 * @var array
	 */
	protected $_layoutRelations = array();
	
	/**
	 * Get master slides
	 * 
	 * @return array
	 */
	public function getMasterSlides()
	{
		return $this->_masterSlides;
	}
	
	/**
	 * Get master slide relations
	 * 
	 * @return array
	 */
	public function getMasterSlideRelations()
	{
		return $this->_masterSlideRelations;
	}
	
	/**
	 * Get themes
	 *
	 * @return array
	 */
	public function getThemes()
	{
		return $this->_themes;
	}
	
	/**
	 * Get theme relations
	 *
	 * @return array
	 */
	public function getThemeRelations()
	{
		return $this->_themeRelations;
	}
	
	/**
	 * Get array of slide layouts
	 * 
	 * @return array
	 */
	public function getLayouts()
	{
		return $this->_layouts;
	}
	
	/**
	 * Get array of slide layout relations
	 * 
	 * @return array
	 */
	public function getLayoutRelations()
	{
		return $this->_layoutRelations;
	}
	
	/**
	 * Find specific slide layout.
	 * 
	 * This is an array consisting of:
	 * - masterid
	 * - name (string)
	 * - body (string)
	 * 
	 * @return array
	 * @throws Exception
	 */
	public function findLayout($name = '', $masterId = 1)
	{
		foreach ($this->_layouts as $layout)
		{
			if ($layout['name'] == $name && $layout['masterid'] == $masterId)
			{
				return $layout;
			}
		}
		
		throw new Exception("Could not find slide layout $name in current layout pack.");
	}
	
	/**
	 * Find specific slide layout index.
	 * 
	 * @return int
	 * @throws Exception
	 */
	public function findLayoutIndex($name = '', $masterId = 1)
	{
		$i = 0;
		foreach ($this->_layouts as $layout)
		{
			if ($layout['name'] == $name && $layout['masterid'] == $masterId)
			{
				return $i;
			}
			
			++$i;
		}
		
		throw new Exception("Could not find slide layout $name in current layout pack.");
	}
}
