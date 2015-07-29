# Changelog
## 0.5.0 - Not released

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
