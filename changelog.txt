**************************************************************************************
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
* @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
* @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
* @version    ##VERSION##, ##DATE##
**************************************************************************************


To be planned:
- General:  (MB) CP-1165 - Rename PHPPowerpoint.php to PHPPowerPoint.php
- General:  (MB) CP-5270 - Create build script using Phing
- Feature:  (MB) CP- 815 - Provide fluent interfaces where possible
- Feature:  (MB) CP-1034 - Use existing presentation template when writing PPTX file
- Feature:  (MB) CP-1093 - Implement bullet and numeric lists
- Feature:  (MB) CP-1173 - getProperties: setCompany feature request
- Feature:  (MB) CP-1375 - New shape type: table
- Feature:  (MB) CP-2804 - Use of CDATA text when writing text
- Feature:  (MB) CP-1378 - Possibility to set borders on tables and table cells
- Feature:  (MB) CP-4921 - Access to additional properties of Text Boxes
- Feature:  (MB) CP-7010 - Applied patch 7010
- Feature:  (MB) CP-7020 - Applied patch 7020
- Feature:  (MB) CP-1196 - Add a hyperlink to an image or textbox
- Feature:  (MBaker) Implement autoloader
- Feature:  (Progi1984) Implement ODP Writer
- Feature:  (MB) CP-4953 - PowerPoint Charts
- Feature:  (MB) CP-5580 - Editing chart data
- Feature:  (MB) CP-5461 - Solid Fill support
- Feature:  (MB) CP-8375 - Applied patch 8375
- Feature:  (Progi1984) Implement ODP Writer
- Bugfix:   (MBaker) Allow solid colour fill
- Bugfix:   (MB) CP-3910 - Table width setting Office 2007
- Bugfix:   (MB) CP-4598 - Bullet characters in Master Slide Layouts of template file become corrupted
- Bugfix:   (MB) CP-3424 - Generated files cannot be opened in Office 08 for Mac OSX
- Bugfix:   (MB) CP-2541 - Table Cell Borders Not Displaying Correctly
- Bugfix:   (MB) CP-4597 - Multiple Master Slides are not supported 
- Bugfix:   (MB) CP-4596 - Images in Layouts other than first Master Slide within Template file causes corrupted PPTX
- Bugfix :  (delphiki) GH-16 - Fixed A3 and A4 formats dimensions 
- Bugfix :  (delphiki) GH-18 - Fixed custom document layout

Initial version:
- Create a Presentation object
- Add one or more Slide objects
- Add one or more Shapes to Slide objects
- Text Shapes
- Image Shapes
- Export Presentation object to PowerPoint 2007 OpenXML format
