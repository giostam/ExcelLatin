# ExcelLatin
Transliterate excel (xlsx) columns from Greek to Latin and vice versa

usage: java -jar ExcelLatin [options] filenameIn filenameOut columnNames...

options:
---
-L to Latin

-G to Greek

-d do not deaccent the result

-h help

You can create a map.txt file in the same path as the ExcelLatin.jar. The file will have values that will skip the transliteration and will be directly mapped. It should have two tab separated values in each row, first the Greek value and then the Latin one.

### Examples
*java -jar ExcelLatin.jar a.xlsx b.xlsx column1 column2*

It will transliterate the columns "column1" and "column2" of the a.xlsx excel file from Greek to Latin, except for the values that are in the map.txt (if it exists)



*java -jar ExcelLatin.jar -G -d a.xlsx b.xlsx column1 column2*

It will transliterate the columns "column1" and "column2" of the a.xlsx excel file from Latin to Greek without deaccenting the resulting value, except for the values that are in the map.txt (if it exists)

Based on icu4j transliterator: http://userguide.icu-project.org/
