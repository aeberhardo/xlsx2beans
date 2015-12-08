xlsx2beans
==========

Convert XLSX and XLS spreadsheets to Java beans.

Requirements
------------
* [Apache POI 3.9](http://poi.apache.org/)
* Java 7

Features
--------
* Annotation `@XlsxColumnName` for mapping spreadsheet columns to Java bean properties
* Supports Strings, Numerics (Numbers and Dates) and Formulas

History
-------
0.0.7
Merged enhancements by [azuxx](https://github.com/azuxx):
* Supported conversion of XLS files as well
* Supported annotation `@XlsxColumnName` on bean super class if exists
