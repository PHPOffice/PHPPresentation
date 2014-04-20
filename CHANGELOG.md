# Changelog

## 0.2.0 - Not yet released

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
- Implement ODP Writer - @Progi1984

### Bugfix

- Allow solid colour fill - @MarkBaker
- Table width setting Office 2007 - @maartenba CP-3910
- Bullet characters in Master Slide Layouts of template file become corrupted - @maartenba CP-4598
- Generated files cannot be opened in Office 08 for Mac OSX - @maartenba CP-3424
- Table Cell Borders Not Displaying Correctly - @maartenba CP-2541
- Multiple Master Slides are not supported - @maartenba CP-4597
- Images in Layouts other than first Master Slide within Template file causes corrupted PPTX - @maartenba CP-4596
- Fixed A3 and A4 formats dimensions - @delphiki GH-16
- Fixed custom document layout - @delphiki GH-18

### Miscellaneous

- Rename PHPPowerpoint.php to PHPPowerPoint.php - @maartenba CP-1165
- Create build script using Phing - @maartenba CP-5270
- QA: Prepare `.travis.yml` and `phpcs.xml` for Travis build passing - @Progi1984 @ivanlanin
- QA: Initiate unit tests - @Progi1984 @ivanlanin
- QA: Cleanup source code for PSR dan PHPDoc compatibility - @ivanlanin
- Doc: Initiate documentation - @ivanlanin
- Refactor: Change PHPPowerPoint_Shape_Shadow to PHPPowerPoint_Style_Shadow because it's a style, not a shape - @ivanlanin
- Refactor: Change PHPPowerPoint_SlideIterator to PHPPowerPoint_Slide_Iterator - @ivanlanin

## 0.1.0

- Create a Presentation object
- Add one or more Slide objects
- Add one or more Shapes to Slide objects
- Text Shapes
- Image Shapes
- Export Presentation object to PowerPoint 2007 OpenXML format
