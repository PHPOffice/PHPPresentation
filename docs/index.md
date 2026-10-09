#

![PHPPresentation](images/PHPPresentationLogo.png)

PHPPresentation is a library written in pure PHP that provides a set of
classes to write to different presentation file formats, i.e.
[Microsoft Office Open XML](http://en.wikipedia.org/wiki/Office_Open_XML)
(`.pptx`) and OASIS [Open Document Format for Office Applications](http://en.wikipedia.org/wiki/OpenDocument) (`.odp`).

PHPPresentation is an open source project licensed under the terms of [LGPL
version 3](https://github.com/PHPOffice/PHPPresentation/blob/develop/COPYING.LESSER).
PHPPresentation is aimed to be a high quality software product by incorporating
[continuous integration and unit testing](https://github.com/PHPOffice/PHPPresentation/actions/workflows/php.yml).
You can learn more about PHPPresentation by reading this Developers'
Documentation.
<!---
-  and the `API Documentation <http://phpoffice.github.io/PHPPresentation/docs/develop/>`__
-->

## Features

- Create an in-memory presentation representation
- Set presentation meta data (author, title, description, etc)
- Add slides from scratch or from existing one
- Supports different fonts and font styles
- Supports different formatting, styles, fills, gradients
- Supports hyperlinks and rich-text strings
- Add images with different styles (positioning, rotation, shadow)
- Set printing options (header, footer, page margins, paper size, orientation)
- Input from different file formats:
    - PowerPoint 97 (.ppt)
    - PowerPoint 2007 (.pptx)
    - OpenDocument Presentation (.odp)
    - Serialized Spreadsheet
- Output to different file formats:
    - PowerPoint 2007 (.pptx)
    - OpenDocument Presentation (.odp)
    - HTML (.html)
    - PDF (.pdf)
    - Serialized Spreadsheet
- ... and lots of other things!

## File formats

Below are the supported features for each file formats. :material-alert-outline: means partly supported: the [Writers](usage/writers.md#supported-features) and [Readers](usage/readers.md#supported-features) pages say what is missing, feature by feature.


### Writers

| Features                  |                      | HTML  | Keynote | ODP   | PDF   | PPTX  | Serialized |
|---------------------------|----------------------|-------|---------|-------|-------|-------|------------|
| **Document**              | Mark as final        |       |         |       |       | :material-check: | :material-check: |
| **Document Properties**   | Standard             |       |         | :material-alert-outline: |       | :material-check: | :material-check: |
|                           | Custom               |       |         | :material-check: |       | :material-check: | :material-check: |
| **Slides**                |                      | :material-alert-outline: | :material-check: | :material-check: | :material-check: | :material-check: | :material-check: |
|                           | Name                 |       |         | :material-check: |       | :material-check: | :material-check: |
| **Element Shape**         | AutoShape            |       |         | :material-check: |       | :material-check: | :material-check: |
|                           | Image                | :material-check: | :material-check: | :material-check: | :material-check: | :material-check: | :material-alert-outline: |
|                           | Hyperlink            | :material-check: |         | :material-check: | :material-check: | :material-check: | :material-check: |
|                           | Line                 |       |         | :material-check: |       | :material-check: | :material-check: |
|                           | MemoryImage          | :material-check: | :material-check: | :material-check: | :material-check: | :material-check: | :material-alert-outline: |
|                           | RichText             | :material-check: | :material-alert-outline: | :material-check: | :material-check: | :material-check: | :material-check: |
|                           | Table                | :material-alert-outline: |         | :material-check: | :material-alert-outline: | :material-check: | :material-check: |
|                           | Text                 | :material-alert-outline: |         | :material-check: | :material-alert-outline: | :material-check: | :material-check: |
|                           | Group                |       |         | :material-check: |       | :material-check: | :material-alert-outline: |
|                           | Media                | :material-alert-outline: |         | :material-check: |       | :material-check: | :material-check: |
|                           | Comment              |       |         | :material-check: |       | :material-check: | :material-check: |
| **Charts**                | Area                 |       |         | :material-check: |       | :material-check: | :material-check: |
|                           | Bar                  |       |         | :material-alert-outline: |       | :material-check: | :material-check: |
|                           | Bar3D                |       |         | :material-alert-outline: |       | :material-check: | :material-check: |
|                           | Doughnut             |       |         | :material-alert-outline: |       | :material-check: | :material-check: |
|                           | Line                 |       |         | :material-check: |       | :material-check: | :material-check: |
|                           | Pie                  |       |         | :material-check: |       | :material-check: | :material-check: |
|                           | Pie3D                |       |         | :material-alert-outline: |       | :material-check: | :material-check: |
|                           | Radar                |       |         | :material-check: |       | :material-check: | :material-check: |
|                           | Scatter              |       |         | :material-check: |       | :material-check: | :material-check: |


### Readers

| Features                  |                      | Keynote | ODP   | PPT   | PPTX  | Serialized |
|---------------------------|----------------------|---------|-------|-------|-------|------------|
| **Document**              | Mark as final        |         |       |       | :material-check: | :material-check: |
| **Document Properties**   | Standard             |         | :material-alert-outline: |       | :material-alert-outline: | :material-check: |
|                           | Custom               |         | :material-check: |       | :material-check: | :material-check: |
| **Slides**                |                      | :material-alert-outline: | :material-check: | :material-check: | :material-check: | :material-check: |
|                           | Name                 |         | :material-check: |       | :material-check: | :material-check: |
| **Element Shape**         | AutoShape            |         | :material-alert-outline: |       | :material-alert-outline: | :material-check: |
|                           | Image                | :material-alert-outline: | :material-check: | :material-alert-outline: | :material-check: | :material-alert-outline: |
|                           | Hyperlink            |         | :material-check: | :material-alert-outline: | :material-check: | :material-check: |
|                           | Line                 |         | :material-alert-outline: | :material-alert-outline: | :material-alert-outline: | :material-check: |
|                           | MemoryImage          | :material-alert-outline: | :material-check: | :material-alert-outline: | :material-check: | :material-alert-outline: |
|                           | RichText             | :material-alert-outline: | :material-check: | :material-check: | :material-check: | :material-check: |
|                           | Table                |         | :material-check: |       | :material-check: | :material-check: |
|                           | Text                 |         | :material-check: | :material-alert-outline: | :material-check: | :material-check: |
|                           | Group                |         | :material-check: | :material-alert-outline: | :material-check: | :material-alert-outline: |
|                           | Media                |         |       |       |       | :material-check: |
|                           | Comment              |         |       |       |       | :material-check: |
| **Charts**                | Area                 |         | :material-check: |       | :material-check: | :material-check: |
|                           | Bar                  |         | :material-alert-outline: |       | :material-check: | :material-check: |
|                           | Bar3D                |         | :material-alert-outline: |       | :material-alert-outline: | :material-check: |
|                           | Doughnut             |         | :material-alert-outline: |       | :material-check: | :material-check: |
|                           | Line                 |         | :material-check: |       | :material-check: | :material-check: |
|                           | Pie                  |         | :material-check: |       | :material-check: | :material-check: |
|                           | Pie3D                |         | :material-alert-outline: |       | :material-alert-outline: | :material-check: |
|                           | Radar                |         | :material-check: |       | :material-check: | :material-check: |
|                           | Scatter              |         | :material-check: |       | :material-check: | :material-check: |


## Contributing

We welcome everyone to contribute to PHPPresentation. Below are some of the things that you can do to contribute:

-  [Fork us](https://github.com/PHPOffice/PHPPresentation/fork) and [request a pull](https://github.com/PHPOffice/PHPPresentation/pulls) to the [master](https://github.com/PHPOffice/PHPPresentation) branch
-  Submit [bug reports or feature requests](https://github.com/PHPOffice/PHPPresentation/issues) to GitHub
