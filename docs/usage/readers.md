# Readers

## Keynote
The name of the reader is `Keynote`.

``` php
<?php

$reader = IOFactory::createReader('Keynote');
$reader->load(__DIR__ . '/sample.key');
```

The reader reads the text of every slide, its speaker note and the images the slide uses, out of a
Keynote '13 (and later) package -- its `Index/*.iwa` components -- as well as out of the
`index.apxl` document of a Keynote '09 package. Anything else the format carries (styles, tables,
charts, transitions) is not read yet.

### Options

#### Load without images

You can load a presentation without images.

``` php
<?php

use PhpOffice\PhpPresentation\Reader\Keynote;

$reader = new Keynote();
$reader->load(__DIR__ . '/sample.key', Keynote::SKIP_IMAGES);
```

## ODPresentation
The name of the reader is `ODPresentation`.

``` php
<?php

$reader = IOFactory::createReader('ODPresentation');
$reader->load(__DIR__ . '/sample.odp');
```

### Options

#### Load without images

You can load a presentation without images.

``` php
<?php

use PhpOffice\PhpPresentation\Reader\ODPresentation;

$reader = new ODPresentation();
$reader->load(__DIR__ . '/sample.odp', ODPresentation::SKIP_IMAGES);
```

## PowerPoint97
The name of the reader is `PowerPoint97`.

``` php
<?php

$reader = IOFactory::createReader('PowerPoint97');
$reader->load(__DIR__ . '/sample.ppt');
```

### Options

#### Load without images

You can load a presentation without images.

``` php
<?php

use PhpOffice\PhpPresentation\Reader\PowerPoint97;

$reader = new PowerPoint97();
$reader->load(__DIR__ . '/sample.ppt', PowerPoint97::SKIP_IMAGES);
```

## PowerPoint2007
The name of the reader is `PowerPoint2007`.

``` php
<?php

$reader = IOFactory::createReader('PowerPoint2007');
$reader->load(__DIR__ . '/sample.pptx');
```

### Options

#### Load without images

You can load a presentation without images.

``` php
<?php

use PhpOffice\PhpPresentation\Reader\PowerPoint2007;

$reader = new PowerPoint2007();
$reader->load(__DIR__ . '/sample.pptx', PowerPoint2007::SKIP_IMAGES);
```

## Serialized
The name of the reader is `Serialized`.

``` php
<?php

$reader = IOFactory::createReader('Serialized');
$reader->load(__DIR__ . '/sample.phppt');
```

## Supported features

Keynote reads Keynote '09 (APXL) and Keynote '13+ (IWA). WMF and EMF (EMF+ too) images need
`phpoffice/wmf` (a suggested dependency) and GD: the PPTX and ODP Readers keep them as they are, and the
Keynote Reader reads them. Serialized reads back the whole object graph, so it has every feature of the
model, with the exceptions its notes list.

:material-check: supported, :material-alert-outline: partly supported, :material-close: not supported, empty: the format has no such feature. A number points to its note under the table.

| Features |  | Keynote | ODP | PPT | PPTX | Serialized |
| --- | --- | --- | --- | --- | --- | --- |
| **Document** | Standard properties (title, author, subject, keywords, dates...) | :material-close: | :material-alert-outline:<sup>1</sup> | :material-close: | :material-alert-outline:<sup>2</sup> | :material-check: |
|  | Custom properties |  | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Mark as final |  |  | :material-close: | :material-check: | :material-check: |
|  | Slide size / layout | :material-close: | :material-close: | :material-close: | :material-check: | :material-check: |
|  | Default language | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Presentation properties (slideshow type, loop, zoom, last view) | :material-close: | :material-alert-outline:<sup>3</sup> | :material-close: | :material-alert-outline:<sup>4</sup> | :material-check: |
| **Slides** | Slides | :material-alert-outline:<sup>5</sup> | :material-check: | :material-check: | :material-check: | :material-check: |
|  | Slide name |  | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Hidden slide | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Background (colour, image) | :material-close: | :material-check: | :material-close: | :material-check: | :material-alert-outline:<sup>6</sup> |
|  | Transition | :material-close: | :material-close: | :material-close: | :material-close: | :material-check: |
|  | Animation | :material-close: | :material-close: | :material-close: | :material-close: | :material-check: |
|  | Speaker notes | :material-alert-outline:<sup>7</sup> | :material-check: | :material-alert-outline:<sup>8</sup> | :material-check: | :material-check: |
|  | Masters and layouts | :material-close: | :material-close: | :material-close: | :material-check: | :material-alert-outline:<sup>9</sup> |
|  | Placeholders (title, body, date, footer, slide number) | :material-close: | :material-check: | :material-close: | :material-alert-outline:<sup>10</sup> | :material-check: |
| **Shape (all)** | Name |  | :material-alert-outline:<sup>11</sup> | :material-close: | :material-check: | :material-check: |
|  | Description (alternative text) | :material-close: | :material-alert-outline:<sup>12</sup> | :material-close: | :material-check: | :material-check: |
|  | Decorative flag |  | :material-check: |  | :material-check: | :material-check: |
|  | Hyperlink on the shape | :material-close: | :material-close: | :material-close: | :material-check: | :material-check: |
|  | Hyperlink tooltip |  | :material-close: | :material-close: | :material-check: | :material-check: |
|  | Position and size | :material-alert-outline:<sup>13</sup> | :material-check: | :material-check: | :material-check: | :material-check: |
|  | Rotation | :material-close: | :material-check: | :material-alert-outline:<sup>14</sup> | :material-alert-outline:<sup>15</sup> | :material-check: |
|  | Fill | :material-close: | :material-alert-outline:<sup>16</sup> | :material-close: | :material-alert-outline:<sup>17</sup> | :material-check: |
|  | Border / outline | :material-close: | :material-alert-outline:<sup>18</sup> | :material-alert-outline:<sup>19</sup> | :material-alert-outline:<sup>20</sup> | :material-check: |
|  | Shadow | :material-close: | :material-alert-outline:<sup>21</sup> | :material-alert-outline:<sup>22</sup> | :material-alert-outline:<sup>23</sup> | :material-check: |
|  | Order (z-order) of shapes on a slide | :material-alert-outline:<sup>24</sup> | :material-check: | :material-check: | :material-check: | :material-check: |
| **Shape types** | AutoShape | :material-close: | :material-alert-outline:<sup>25</sup> | :material-close: | :material-alert-outline:<sup>26</sup> | :material-check: |
|  | Image (file, Base64, ZIP, GD / memory) | :material-alert-outline:<sup>27</sup> | :material-check: | :material-alert-outline:<sup>28</sup> | :material-check: | :material-alert-outline:<sup>29</sup> |
|  | Media (audio, video) | :material-close: | :material-close: | :material-close: | :material-close: | :material-check: |
|  | Line | :material-close: | :material-alert-outline:<sup>30</sup> | :material-alert-outline:<sup>31</sup> | :material-alert-outline:<sup>32</sup> | :material-check: |
|  | Group | :material-close: | :material-check: | :material-alert-outline:<sup>33</sup> | :material-check: | :material-alert-outline:<sup>34</sup> |
|  | RichText (text box) | :material-alert-outline:<sup>35</sup> | :material-check: | :material-check: | :material-check: | :material-check: |
|  | Table | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Chart | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Comment | :material-close: | :material-close: | :material-close: | :material-close: | :material-check: |
| **Text** | Paragraph alignment and indent | :material-close: | :material-alert-outline:<sup>36</sup> | :material-alert-outline:<sup>37</sup> | :material-check: | :material-check: |
|  | Line spacing, spacing before/after | :material-close: | :material-alert-outline:<sup>38</sup> | :material-close: | :material-check: | :material-check: |
|  | Bullets (character, colour, size) | :material-close: | :material-alert-outline:<sup>39</sup> | :material-alert-outline:<sup>40</sup> | :material-check: | :material-check: |
|  | Numbered lists (format, start) | :material-close: | :material-close: | :material-close: | :material-check: | :material-check: |
|  | Font: name, size, bold, italic, colour | :material-close: | :material-check: | :material-alert-outline:<sup>41</sup> | :material-check: | :material-check: |
|  | Font: underline, strikethrough, superscript/subscript, caps | :material-close: | :material-alert-outline:<sup>42</sup> | :material-alert-outline:<sup>43</sup> | :material-check: | :material-check: |
|  | Language of a run | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Hyperlink on text | :material-close: | :material-check: | :material-alert-outline:<sup>44</sup> | :material-check: | :material-check: |
|  | Fields (date/time, slide number) | :material-close: | :material-alert-outline:<sup>45</sup> | :material-close: | :material-check: | :material-check: |
|  | Text box: insets, autofit, wrap, columns, vertical/RTL | :material-close: | :material-alert-outline:<sup>46</sup> | :material-alert-outline:<sup>47</sup> | :material-alert-outline:<sup>48</sup> | :material-check: |
| **Table** | Merged cells (column span, row span) | :material-close: | :material-close: | :material-close: | :material-check: | :material-check: |
|  | Cell borders and fill | :material-close: | :material-alert-outline:<sup>49</sup> | :material-close: | :material-check: | :material-check: |
|  | Header row | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Column widths, row heights | :material-close: | :material-alert-outline:<sup>50</sup> | :material-close: | :material-check: | :material-check: |
| **Charts** | Area | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Bar | :material-close: | :material-alert-outline:<sup>51</sup> | :material-close: | :material-check: | :material-check: |
|  | Bar3D | :material-close: | :material-alert-outline:<sup>52</sup> | :material-close: | :material-alert-outline:<sup>53</sup> | :material-check: |
|  | Doughnut | :material-close: | :material-alert-outline:<sup>54</sup> | :material-close: | :material-check: | :material-check: |
|  | Line | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Pie | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Pie3D | :material-close: | :material-alert-outline:<sup>55</sup> | :material-close: | :material-alert-outline:<sup>56</sup> | :material-check: |
|  | Radar | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Scatter | :material-close: | :material-check: | :material-close: | :material-check: | :material-check: |
|  | Chart title | :material-close: | :material-check: | :material-close: | :material-alert-outline:<sup>57</sup> | :material-check: |
|  | Legend | :material-close: | :material-alert-outline:<sup>58</sup> | :material-close: | :material-alert-outline:<sup>59</sup> | :material-check: |
|  | Axes (title, min/max, gridlines, format) | :material-close: | :material-alert-outline:<sup>60</sup> | :material-close: | :material-alert-outline:<sup>61</sup> | :material-check: |
|  | Series: labels, colours, markers | :material-close: | :material-alert-outline:<sup>62</sup> | :material-close: | :material-alert-outline:<sup>63</sup> | :material-check: |
|  | Chart language |  | :material-close: | :material-close: | :material-close: | :material-check: |

1. Category, company, revision, status not read; dates only with zone, no fraction.
2. Company (app.xml) not read.
3. Browse type only; loop not read.
4. Last view, comment visibility not read.
5. IWA order from component names, not slide tree.
6. Background image kept as external path.
7. Plain text only.
8. First shape only; breaks with several slides.
9. Master/layout images kept as external paths.
10. Type only; idx, untyped ph dropped.
11. Images, charts, AutoShapes only.
12. Falls back to name when absent.
13. IWA: no geometry read, APXL only.
14. Raw 16.16 value, not converted to degrees.
15. Not charts or tables.
16. None and solid only.
17. Chart fill not read.
18. Not lines.
19. Line colour only; shape outlines ignored.
20. Not lines or charts.
21. Not lines.
22. Offset only, direction fixed 45°, no colour.
23. Not lines, groups or charts.
24. IWA: images placed after all text.
25. Rounded corner not read.
26. Rectangles come back as RichText.
27. Raster and WMF/EMF only; SVG/PDF images dropped.
28. JPEG/PNG blips only; others throw.
29. Base64 and GD not loadable.
30. Ends only; line style not read.
31. Height overwritten by line width; no flip.
32. Ends only; line style not read.
33. Group geometry not read; child anchors raw.
34. Child images kept as external paths.
35. Plain text; formatting not read.
36. List items lose alignment; hanging indent read positive.
37. Alignment applied to first paragraph only.
38. Lost on list items.
39. No bullet colour.
40. Character only; no bullet colour, size, font.
41. Font name keeps NUL padding.
42. Small caps read from lowercase.
43. Single underline only; super/subscript throws.
44. External URLs only.
45. Date/time format not read.
46. No autofit, no vertical text.
47. Insets only; no autofit/wrap/columns/direction.
48. No autofit, columns, vertical text.
49. Cell fill solid only.
50. Row heights only.
51. Gap width and overlap not read.
52. View3D, gap width not read.
53. View3D not read.
54. Hole size, first slice angle not read.
55. View3D not read.
56. View3D not read.
57. Title position not read.
58. No fill or border.
59. Visibility and position only.
60. No number format, tick marks, reverse order.
61. No min/max, units, gridlines, number format.
62. No label position or series-name label.
63. No markers, label position, font.
