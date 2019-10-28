# pml.xsd

## How do I create the file ?
* Go to http://www.ecma-international.org/publications/standards/Ecma-376.htm
* Download the file "ECMA-376 1st edition Part 4"
* Extract the directory Office Open XML 1st edition Part 2 (DOCX).zip\OpenPackagingConventions-XMLSchema.zip
* From this, create files (dml-diagram.xsd, dml-main.xsd, pml.xsd, sml.xsd) 
  * Incorporate xsd files based on the order of **ECMA-376, Second Edition, Part 4 - Transitional Migration Features.zip\OfficeOpenXML-XMLSchema-Transitional.zip** files
  * Try to avoid duplicate imprts

## How do valid new created xsd ?
* Fetch some files from an PPTX files (theme1.xml, ...)
* Use xmllint :
```
$ xmllint --noout --schema pml.xsd ../../../../samples/results/theme1.xml
```