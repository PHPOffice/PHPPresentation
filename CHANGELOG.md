# Changelog

## 0.9.0 - 2017-07-05

### Bugfix
- PowerPoint2007 Writer : Margins in table cell - @Progi1984 GH-347

### Changes
- PowerPoint2007 Writer : Write percentage values with a trailing percent sign instead of formatted as 1000th of a percent to comply with the standard - @k42b3 GH-307

### Features
- PowerPoint2007 Writer : Implemented XSD validation test case according to the ECMA/ISO standard - @k42b3 GH-307
- PowerPoint2007 Writer : Implement visibility for axis - @kw-pr @Progi1984 GH-356 
- PowerPoint2007 Writer : Implement gap width in Bar(3D) Charts - @Progi1984 GH-358

## 0.8.0 - 2017-04-03

### Bugfix
- PowerPoint2007 Writer : Fixed the marker on line chart when symbol is none - @Napryc GH-211
- PowerPoint2007 Writer : Fixed the format value in Pie Chart - @Napryc GH-212
- PowerPoint2007 Writer : The presentation need repairs on Mac @jrking4 GH-266 GH-276
- PowerPoint2007 Writer : Fix for PowerPoint2007 Writer (Need repair) @Progi1984 GH-266 GH-274 GH-276 GH-282 GH-302
- PowerPoint2007 Writer : Fixed the axis title in bar chart - @pgee70 GH-267
- PowerPoint2007 Writer : Fixed the label position in bar chart - @pgee70 GH-268
- PowerPoint2007 Writer : Support of margins in cell in table - @carlosafonso @Progi1984 GH-273 GH-315
- Fixed the corruption of file when an addExternalSlide is called - @Progi1984 GH-240

### Changes
- Misc : Added two methods for setting Border & Fill in Legend - @Progi1984 GH-265

### Features
- ODPresentation Writer : Show/Hide Value / Name / Series Name in Chart - @Progi1984 GH-272
- ODPresentation Writer : Axis Bounds in Chart - @Progi1984 GH-269
- PowerPoint97 Reader : Support of Slide Note - @Progi1984 GH-226
- PowerPoint2007 Reader : Support of Shape Table - @Progi1984 GH-240
- PowerPoint2007 Reader : Support of Slide Note - @Progi1984 GH-226
- PowerPoint2007 Reader : Support text direction in Alignment for Table - @Progi1984 GH-218
- PowerPoint2007 Writer : Implement character spacing - @jvanoostrom GH-301
- PowerPoint2007 Writer : Axis Bounds in Chart - @Progi1984 GH-269
- PowerPoint2007 Writer : Implement Legend Key in Series for Chart - @Progi1984 GH-319
- PowerPoint2007 Writer : Support text direction in Alignment for Table - @SeregPie GH-218
- PowerPoint2007 Writer : Support tick mark & unit in Axis for Chart - @Faab GH-218
- PowerPoint2007 Writer : Support separator in Series for Chart - @jphchaput GH-218
- PowerPoint2007 Writer : Add support for Outline in Axis - @Progi1984 GH-255
- PowerPoint2007 Writer : Support autoscale for Chart - @Progi1984 GH-293

## 0.7.0 - 2016-09-12

### Bugfix
- Fixed the image project - @mvargasmoran GH-177
- PowerPoint2007 Writer : Bugfix for printing slide notes - @JewrassicPark @Progi1984 GH-179

### Changes
- PhpOffice\PhpPresentation\Writer\ODPresentation : Move to Design Pattern Decorator - @Progi1984
- PhpOffice\PhpPresentation\Writer\PowerPoint2007 : Move to Design Pattern Decorator - @Progi1984
- PhpOffice\PhpPresentation\Shape\Type\AbstracType\getData has been deprecated for PhpOffice\PhpPresentation\Shape\Type\AbstracType\getSeries - @Progi1984 GH-169
- PhpOffice\PhpPresentation\Shape\Type\AbstracType\setData has been deprecated for PhpOffice\PhpPresentation\Shape\Type\AbstracType\setSeries - @Progi1984 GH-169
- Added documentation for chart series (font, outline, marker) - @Progi1984 GH-169
- Internal Structure for Drawing Shape - @Progi1984 GH-192
- Documentation about manual installation - @danielbair GH-254

### Features
- ODPresentation & PowerPoint2007 Writer : Add support for Comment - @Progi1984 GH-116
- ODPresentation & PowerPoint2007 Writer : Thumbnail of the presentation - @Progi1984 GH-125
- ODPresentation & PowerPoint2007 Writer : Add support for Gridlines in Chart - @Progi1984 GH-129
- ODPresentation & PowerPoint2007 Writer : Support for images in base 64 - @Progi1984 GH-168
- ODPresentation & PowerPoint2007 Writer : Marker of Series in Line & Scatter chart is customizable - @Progi1984 GH-169
- ODPresentation & PowerPoint2007 Writer : Outline of Series in Line & Scatter chart is customizable - @Progi1984 GH-169
- ODPresentation & PowerPoint2007 & Serialized Writer : Support for Zip Adapter - @Progi1984 GH-176
- ODPresentation & PowerPoint2007 Writer : language property to TextElement - @skrajewski & @Progi1984 GH-180
- ODPresentation & PowerPoint2007 Writer : Add Font Support For Chart Axis - @jrking4 GH-186
- ODPresentation & PowerPoint2007 Writer : Support for video - @Progi1984 GH-123
- ODPresentation & PowerPoint2007 Writer : Support for Visibility for slides - @Progi1984
- PowerPoint2007 Reader : Layout Management - @vincentKool @Progi1984 GH-161 
- PowerPoint2007 Reader : Slide size - @loverslcn @Progi1984 GH-246 
- PowerPoint2007 Reader : Bullet Color - @Progi1984 GH-257 
- PowerPoint2007 Reader : Line Spacing - @Progi1984 GH-257 
- PowerPoint2007 Writer : Presentation with predefined View Type - @Progi1984 GH-120
- PowerPoint2007 Writer : Implement alpha channel to Fills - @Dayjo GH-203 / @Progi1984 GH-215
- PowerPoint2007 Writer : Implement Animations - @JewrassicPark GH-214 / @Progi1984 GH-217
- PowerPoint2007 Writer : Layout Management - @vincentKool @Progi1984 GH-161
- PowerPoint2007 Writer : Bullet Color - @piotrbelina GH-249 
- PowerPoint2007 Writer : Line Spacing - @piotrbelina GH-249 

## 0.6.0 - 2016-01-24

### Bugfix
- Documentation : Fixes in the rename of PHPPowerPoint - @Progi1984 GH-127
- ODPresentation : Exclude SVM files for reader - @Progi1984 GH-141
- PowerPoint2007 Writer : Bugfix for opening PPTX on Mac - @thsteinmetz GH-89

### Changes
- PhpOffice\PhpPresentation\getProperties has been deprecated for PhpOffice\PhpPresentation\getDocumentProperties - @Progi1984 GH-154
- PhpOffice\PhpPresentation\setProperties has been deprecated for PhpOffice\PhpPresentation\setDocumentProperties - @Progi1984 GH-154
- PhpOffice\PhpPowerpoint\Style\Alignment::setLevel can now be defined great than 8 - @Progi1984 GH-141

### Features
- ODPresentation Reader/Writer : Name of the slide - @Progi1984 GH-121
- ODPresentation Reader/Writer : Slide Background Color or Image - @Progi1984 GH-152
- PowerPoint2007 Reader : Support for Layout Name - @Progi1984 GH-144
- PowerPoint2007 Reader/Writer : Mark as final - @Progi1984 GH-118
- PowerPoint2007 Reader/Writer : Set default zoom value for presentation - @Progi1984 GH-122
- PowerPoint2007 Reader/Writer : Slide Background Color or Image - @Progi1984 GH-152
- PowerPoint2007 Reader/Writer : Add Properties for allowing loop continuously until 'Esc' - @Progi1984 GH-154

## 0.5.0 - 2015-10-08

### Features
- PowerPoint2007 Reader : Initial Commit - @Progi1984 GH-44
- ODPresentation Reader : Initial Commit - @Progi1984 GH-113

### Bugfix
- Fixed the sample in Readme.md - @Progi1984 GH-114

### Changes
- PhpOffice\PhpPowerpoint becomes PhpOffice\PhpPresentation - @Progi1984 GH-25
- PhpOffice\PhpPowerpoint\Style\Font::setStriketrough has been removed : Use setStrikethrough - @Progi1984
- PhpOffice\PhpPowerpoint\AbstractShape::getSlide has been removed - @Progi1984
- PhpOffice\PhpPowerpoint\AbstractShape::setSlide has been removed - @Progi1984
- PhpOffice\PhpPowerpoint\DocumentLayout::getLayoutXmilli has been removed : getCX(DocumentLayout::UNIT_MILLIMETER) - @Progi1984
- PhpOffice\PhpPowerpoint\DocumentLayout::getLayoutYmilli has been removed : getCY(DocumentLayout::UNIT_MILLIMETER) - @Progi1984
- PhpOffice\PhpPowerpoint\DocumentLayout::setLayoutXmilli has been removed : setCX(DocumentLayout::UNIT_MILLIMETER) - @Progi1984
- PhpOffice\PhpPowerpoint\DocumentLayout::setLayoutYmilli has been removed : setCY(DocumentLayout::UNIT_MILLIMETER) - @Progi1984
- Update the dependence PhpOffice\Common to 0.2.* - @Progi1984
- Migrated Travis CI to legacy - @Progi1984 GH-115

## 0.4.0 - 2015-07-07

### Features
- Added support for grouping shapes together in a Group - @Pr0phet GH-68
- Added support for calculating the offset and extent on a Slide. - @Pr0phet GH-68
- Added support for Horizontal bar chart - @rdoepke @Progi1984 GH-58
- Added support for hyperlink on picture (ODPresentation & PowerPoint2007) - @Progi1984 GH-49
- Added support for hyperlink on richtext (PowerPoint2007) - @JewrassicPark GH-49
- Added support for notes slide (ODPresentation & PowerPoint2007) - @Progi1984 @JewrassicPark GH-63
- Added option for explosion in Pie3D Chart (ODPresentation & PowerPoint2007) - @Progi1984 GH-76
- ODPresentation Writer : Support for fill in RichText - @Progi1984 GH-79
- ODPresentation Writer : Support for border style in RichText - @Progi1984 GH-79
- ODPresentation Writer : Support for Area Chart - @Progi1984 GH-82
- PowerPoint2007 Writer : Support for Area Chart - @Progi1984 GH-82
- ODPresentation Writer : Support for Bar Chart - @Progi1984 GH-82
- PowerPoint2007 Writer : Support for Bar Chart - @Progi1984 GH-82
- Added units in DocumentLayout - @Progi1984 GH-87
- Added support for transitions between slides - @Progi1984
- ODPresentation Writer : Support for Pie Chart & Stack Percent Bar Charts - @jrking4 GH-108
- PowerPoint2007 Writer : Support for Pie Chart & Stack Percent Bar Charts - @jrking4 GH-108

### Bugfix
- PSR-0 via composer broken - @Progi1984 GH-51
- ODPresentation Writer : Title in Legend in chart doesn't displayed - @Progi1984 GH-79
- ODPresentation Writer : Segments in Pie3D Chart are now in clockwise order, as in PowerPoint2007 Writer - @Progi1984 GH-79
- ODPresentation Writer : Axis in Line Chart have not tick marks displayed, as in PowerPoint2007 Writer - @Progi1984 GH-79
- ODPresentation Writer : Shadow don't work for RichTextShapes - @Progi1984 GH-81
- PowerPoint2007 Writer : Fill don't work for RichTextShapes - @Progi1984 GH-61
- PowerPoint2007 Writer : Border don't work for RichTextShapes - @Progi1984 GH-61
- PowerPoint2007 Writer : Hyperlink in table doesn't work - @Progi1984 GH-70
- PowerPoint2007 Writer : AutoFitNormal works with options (fontScale & lineSpacingReduction) - @Progi1984 @desigennaro GH-71
- PowerPoint2007 Writer : Shadow don't work for RichTextShapes - @Progi1984 GH-81
- PowerPoint2007 Writer : Visibility of the Title doesn't work - @Progi1984 GH-107 
- Refactor findLayoutIndex to findLayoutId where it assumes the slideLayout order was sorted. IMPROVED: unit tests - @kenliau GH-95

### Miscellaneous
- Improved the sample 04-Table for having a Text Run in a Cell - @Progi1984 GH-84
- Improved the sample 04-Table for having two links in a Cell - @Progi1984 GH-93
- Improved the documentation about Table Shapes and cell width - @Progi1984 GH-104
- Some parts of code shared between PHPOffice projects have been moved to PhpOffice/Common - @Progi1984
- Refactored the PowerPoint97 Reader for managing the group shape and improving evolutions - @Progi1984 GH-110
- Added a sample (12) for PowerPoint97 Reader with tree of the PhpPowerPoint object - @Progi1984 GH-110

## 0.3.0 - 2014-09-22

### Features
- PowerPoint97 Reader : Implement Basic Reader - @Progi1984 GH-15 GH-14 GH-4
- ODPresentation Writer : Ability to set auto shrink text - @Progi1984 GH-28
- Make package PSR-4 compliant. Autoload classes by composer out of the box - @Djuki GH-41

### Bugfix
- PowerPoint2007 Writer : Powerpoint Repair Error in Office 2010 - @Progi1984 GH-39
- PowerPoint2007 Writer : BUG: Repair Error / Wrong anchor if you don't set vertical alignment different to VERTICAL_BASE - @fregge GH-42
- PowerPoint2007 Writer : Keynote incompatibility - @catrane CP#237322 / @Progi1984 GH-46

### Miscellaneous
- QA : Move AbstractType for Chart - @Progi1984
- QA : Unit Tests - @Progi1984

## 0.2.0 - 2014-07-22

### Features

- Provide fluent interfaces where possible - @maartenba CP- 815
- Use existing presentation template when writing PPTX file - @maartenba CP-1034
- Implement bullet and numeric lists - @maartenba CP-1093
- getProperties: setCompany feature request - @maartenba CP-1173
- New shape type: table - @maartenba CP-1375
- Use of CDATA text when writing text - @maartenba CP-2804
- Possibility to set borders on tables and table cells - @maartenba CP-1378
- Access to additional properties of Text Boxes - @maartenba CP-4921
- Applied patch 7010 - @maartenba CP-7010
- Applied patch 7020 - @maartenba CP-7020
- Add a hyperlink to an image or textbox - @maartenba CP-1196
- PowerPoint Charts - @maartenba CP-4953
- Editing chart data - @maartenba CP-5580
- Solid Fill support - @maartenba CP-5461
- Applied patch 8375 - @maartenba CP-8375
- Implement autoloader - @MarkBaker
- ODPresentation Writer : Implement Basic Writer - @Progi1984 GH-1
- ODPresentation Writer : Implement Support of Charts - @Progi1984 GH-33
- ODPresentation Writer : Implement Support of Lines - @Progi1984 GH-30
- ODPresentation Writer : Implement Support of Tables - @Progi1984 GH-31
- PowerPoint2007 Writer : Implement Support of Fill  - @Progi1984 GH-32

### Bugfix

- Allow solid color fill - @MarkBaker
- Table width setting Office 2007 - @maartenba CP-3910
- Bullet characters in Master Slide Layouts of template file become corrupted - @maartenba CP-4598
- Generated files cannot be opened in Office 08 for Mac OSX - @maartenba CP-3424
- Table Cell Borders Not Displaying Correctly - @maartenba CP-2541
- Multiple Master Slides are not supported - @maartenba CP-4597
- Images in Layouts other than first Master Slide within Template file causes corrupted PPTX - @maartenba CP-4596
- Fixed A3 and A4 formats dimensions - @delphiki GH-16
- Fixed custom document layout - @delphiki GH-18
- Filename parameter is required for IWriter::save method - @sapfeer0k GH-19
- DocumentLayout: Fix incorrect variable assignment - @kaiesh GH-6
- Hyperlink: Wrong input parameter object type in setHyperlink  - @nynka GH-23
- ODPresentation Writer: ODP writer is locale sensitive in the wrong places  - @Progi1984 GH-21
- ODPresentation Writer: Display InMemory Image  - @Progi1984 GH-29
- PowerPoint2007 Writer: Bar3D doesn't display  - @Progi1984 GH-32
- PowerPoint2007 Writer: Changed PowerPoint2007 writer attributes to protected - @delphiki GH-20
- PowerPoint2007 Writer: Scatter chart with numerical X values not working well  - @Progi1984 GH-3
- Shape RichText: Support of Vertical Alignment in PowerPoint2007 - @Progi1984 GH-35


### Miscellaneous

- Rename PHPPowerpoint.php to PHPPowerPoint.php - @maartenba CP-1165
- Create build script using Phing - @maartenba CP-5270
- QA: Prepare `.travis.yml` and `phpcs.xml` for Travis build passing - @Progi1984 @ivanlanin
- QA: Initiate unit tests - @Progi1984 @ivanlanin
- QA: Cleanup source code for PSR dan PHPDoc compatibility - @ivanlanin
- QA: Unit Tests - @Progi1984 & @ivanlanin 
- Doc: Initiate documentation - @ivanlanin
- Doc: Move to [Read The Docs](http://phppowerpoint.readthedocs.org) - @Progi1984
- Refactor: Change PHPPowerPoint_Shape_Shadow to PHPPowerPoint_Style_Shadow because it's a style, not a shape - @ivanlanin
- Refactor: Change PHPPowerPoint_SlideIterator to PHPPowerPoint_Slide_Iterator - @ivanlanin

## 0.1.0

- Create a Presentation object
- Add one or more Slide objects
- Add one or more Shapes to Slide objects
- Text Shapes
- Image Shapes
- Export Presentation object to PowerPoint 2007 OpenXML format
