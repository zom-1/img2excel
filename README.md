img2excel
====
make a Excel mosaic picture from image.

## Description
img2excel.py makes make a Excel mosaic picture from image.

![Original](namiura_small.jpg)
-->
![Excel File](namiura_excel_small.jpg)

```
python img2excel.py  [-c columns ] [-e excel_file] image_file

```

## Installation
```
git clone git://github.com/zom-1/img2excel .
```

## Example
```
python ./img2excel.py namiura.jpg
python ./img2excel.py -c 20 -e test.xlsx namiura.jpg
```

## Requirement
[openpyxl : library to r/w Excel 2010 xlsx/xlsm files](https://openpyxl.readthedocs.org/)

## See also
[img2excel_ascii : make a Excel ascii picture from image](git://github.com/zom-1/img2excelascii)

## Licence
Copyright (c) 2016 zom-1, Released under the
 [MIT](http://opensource.org/licenses/mit-license.php) license

## Author
[zom-1](https://github.com/zom-1)
