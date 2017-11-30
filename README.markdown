# lisp-xl

* Work in progress!

Common Lisp Library for working with  Microsoft Excel XLSX files, hopefully for ETL (extract-transform-load) and data cleansing purposes. This library is geared for working with large files (i.e. 100MB excel files).

It does *not* load the whole file into RAM, thus, processing large files is possible.

## Reason for being

In the business world, you usually want customers to give you data in the form of compressed, tidy CSV files, but they will send you huge XLSX files instead, files that aren't ready for uplaoding... So, read the mundane Microsoft XLSX files using the celestial programming language, Common Lisp!

This lib is going to be oriented to be used on files that are data-dumps, that is: files that have a header row, followed by data rows. In this kind of files, you would expect to have the initial row to be the "header row", and you would expect that the data types on all the rest of the rows are the same (i.e. if column A has to hold number values, there should NOT be a string on cell A115321...)

You can use the (report-cells-type-change) function to take a look at such offending cells. 

## Features

* Does NOT load the whole file into RAM.
* Report of "offending" cells; that is, cells that change data type along the way. 
* Report of unexpected blank rows
* Loads and uncompresses the XLSX sheet file only once, to save time.
* Able to select/load only a range of rows, and/or only a selected range of columns.

## TODO

* Speedups
* Free burger for each user

## Usage

Typical process is as follows:

1. Load the sheet, creating a struct that holds info about the sheet. For this you need the file path, and the name/index of the sheet (1 = first sheet).

2. Query the sheet using the provided functions. 

```common-lisp

;; load lib
(ql:quickload :lisp-xl)

(defparameter *s* (read-sheet *f* 1)) ; read first sheet of file *f*
(process-sheet *s* :max-row 100) ;; obtain rows as cons

;; Do a report on which cells/rows change format along the file.
(report-cells-type-change *s*)

```
## Uses

* ZIP (xlsx files are zipped)
* xmls
* cl-xmlspam (download a copy from my GitHub)

## Installation

For CL newcomers: 
Make sure you have the cl-xmlspam library downloaded; you can copy it to quicklisp's "local-projects" directory. 
Download lisp-xl, copy to "local-projects", then use quicklisp to load the rest of the required libs and compile:

```common-lisp
(ql:quickload :lisp-xl)
```

## Acknowledgements

* Some core xlsx metadata read functions taken from the xlsx library by Carlos Ungil
* Akihide Nano, for putting said library in GitHub

For working with *small* XLSX files in a Cell-oriented way (i.e. "A1", "B4", etc), you might want to take a look at Carlos library:
https://gitlab.common-lisp.net/cungil/xlsx](https://gitlab.common-lisp.net/cungil/xlsx)

Japanese users may want to use Akihide Nano(a-nano) library which is based on Carlos Ungil's:
https://github.com/a-nano/xlsx
This one also includes returning the XLSX file as a-list or p-list.

## License

Licensed under the MIT license.
